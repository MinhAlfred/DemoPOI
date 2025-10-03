package exe2.slideapp.demopoi.controller;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import exe2.slideapp.demopoi.service.PowerPointService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.Parameter;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.Schema;
import io.swagger.v3.oas.annotations.responses.ApiResponse;
import io.swagger.v3.oas.annotations.responses.ApiResponses;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.HashMap;
import java.util.Map;


@RestController
@RequestMapping("/ppt")
@RequiredArgsConstructor
@Tag(name = "PowerPoint Processing", description = "APIs for processing PowerPoint files with text and image placeholders")
@Slf4j
public class PoiController {
    private final PowerPointService powerPointService;
    private final ObjectMapper objectMapper = new ObjectMapper();

    @PostMapping(value = "/process-template", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    @Operation(summary = "Process PowerPoint template with text and images",
               description = """
               Processes a PowerPoint template by replacing:
               - Text placeholders: {key} format (e.g., {name}, {title})
               - Image placeholders: {IMAGE:key} format (e.g., {IMAGE:logo}, {IMAGE:photo})
               
               How to use:
               1. In your PowerPoint template, use {name} for text replacement
               2. Use {IMAGE:logo} where you want to insert an image
               3. Send the template file with data JSON and image files
               """)
    @ApiResponses(value = {
        @ApiResponse(responseCode = "200",
                     description = "PowerPoint file processed successfully",
                     content = @Content(mediaType = "application/vnd.openxmlformats-officedocument.presentationml.presentation")),
        @ApiResponse(responseCode = "400",
                     description = "Invalid file format or request parameters",
                     content = @Content),
        @ApiResponse(responseCode = "500",
                     description = "Internal server error during processing",
                     content = @Content)
    })
    public ResponseEntity<FileSystemResource> processTemplate(
            @Parameter(description = "PowerPoint template file (.pptx)", required = true)
            @RequestPart("file") MultipartFile file,

            @Parameter(description = """
                       JSON string containing key-value pairs for text placeholder replacement.
                       Example: {"name":"John Doe","title":"Senior Manager","company":"Tech Corp"}
                       """, required = true)
            @RequestPart("data") String dataJson,

            @Parameter(description = """
                       Array of image files for image placeholders (optional).
                       Each image will be mapped using the imageMapping parameter.
                       """)
            @RequestPart(value = "images", required = false) MultipartFile[] imageFiles,

            @Parameter(description = """
                       JSON mapping of image placeholder keys to their corresponding array indices (optional).
                       Example: {"logo":0,"photo":1,"signature":2}
                       This maps images[0] to {IMAGE:logo}, images[1] to {IMAGE:photo}, etc.
                       """)
            @RequestPart(value = "imageMapping", required = false) String imageMappingJson) throws Exception {

        log.info("Processing PowerPoint template: {}", file.getOriginalFilename());

        // Parse text data
        Map<String, String> data = objectMapper.readValue(dataJson, new TypeReference<Map<String, String>>() {});
        log.info("Text placeholders to replace: {}", data.keySet());

        // Parse image mapping and create image map
        Map<String, MultipartFile> images = new HashMap<>();
        if (imageFiles != null && imageFiles.length > 0) {
            log.info("Received {} image files", imageFiles.length);

            if (imageMappingJson != null && !imageMappingJson.trim().isEmpty()) {
                Map<String, Integer> imageMapping = objectMapper.readValue(imageMappingJson, new TypeReference<Map<String, Integer>>() {});

                for (Map.Entry<String, Integer> entry : imageMapping.entrySet()) {
                    String key = entry.getKey();
                    Integer index = entry.getValue();

                    if (index >= 0 && index < imageFiles.length) {
                        images.put(key, imageFiles[index]);
                        log.info("Mapped image placeholder '{}' to file: {} ({})",
                                key, imageFiles[index].getOriginalFilename(), formatFileSize(imageFiles[index].getSize()));
                    } else {
                        log.warn("Image index {} for key '{}' is out of bounds (array size: {})",
                                index, key, imageFiles.length);
                    }
                }
            } else {
                log.warn("Images provided but no imageMapping found. Images will be ignored.");
            }
        }

        log.info("Starting template processing: {} text replacements, {} image replacements",
                data.size(), images.size());

        // Process the template
        File processedFile = powerPointService.processTemplate(file, data, images);

        // Prepare response
        FileSystemResource resource = new FileSystemResource(processedFile);
        String outputFilename = "processed_" + file.getOriginalFilename();

        log.info("Template processed successfully. Output file: {}", outputFilename);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + outputFilename + "\"")
                .header(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate")
                .header(HttpHeaders.PRAGMA, "no-cache")
                .header(HttpHeaders.EXPIRES, "0")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.presentationml.presentation"))
                .contentLength(processedFile.length())
                .body(resource);
    }

    /**
     * Format file size for logging
     */
    private String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
        return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
    }
}
