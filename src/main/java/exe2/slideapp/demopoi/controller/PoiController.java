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
import java.util.Map;


@RestController
@RequestMapping("/ppt")
@RequiredArgsConstructor
@Tag(name = "PowerPoint Processing", description = "APIs for processing PowerPoint files")
public class PoiController {
    private final PowerPointService powerPointService;

    @PostMapping(value = "/process-template", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    @Operation(summary = "Process PowerPoint template",
               description = "Processes a PowerPoint template by filling placeholders with provided data and returns the processed file")
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
            @Parameter(description = "JSON string containing key-value pairs for placeholder replacement. Example: {\"name\":\"John\",\"title\":\"Manager\"}", required = true)
            @RequestPart("data") String dataJson) throws Exception {

        ObjectMapper objectMapper = new ObjectMapper();
        Map<String, String> data = objectMapper.readValue(dataJson, new TypeReference<Map<String, String>>() {});

        // Process the template
        File processedFile = powerPointService.processTemplate(file, data);

        // Return the processed file
        FileSystemResource resource = new FileSystemResource(processedFile);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=processed_template.pptx")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.presentationml.presentation"))
                .body(resource);
    }
}
