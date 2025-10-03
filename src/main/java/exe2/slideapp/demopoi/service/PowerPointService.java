package exe2.slideapp.demopoi.service;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.*;
import java.util.List;

@Service
@Slf4j
public class PowerPointService {

    public java.io.File processTemplate(MultipartFile file, Map<String, String> data, Map<String,MultipartFile> images) throws Exception {
        java.io.File tempFile = java.io.File.createTempFile("pptx-template", ".pptx");

        try (InputStream fis = file.getInputStream();
             XMLSlideShow ppt = new XMLSlideShow(fis);
             FileOutputStream fos = new FileOutputStream(tempFile)) {

            // Debug: in ra số lượng slides
            log.info("Total slides: " + ppt.getSlides().size());
            log.info("Processing {} text placeholders and {} image placeholders", data.size(), images.size());

            int slideIndex = 0;
            for (XSLFSlide slide : ppt.getSlides()) {
                log.info("Processing slide " + (slideIndex + 1));
                replaceTextInSlide(slide, data);
                replaceImagesInSlide(slide, images, ppt);
                slideIndex++;
            }

            ppt.write(fos);
        }

        return tempFile;
    }

    /**
     * Duyệt toàn bộ shape trong 1 slide và thay thế placeholder {KEY} bằng value
     */
    private void replaceTextInSlide(XSLFSlide slide, Map<String, String> data) {
        for (XSLFShape shape : slide.getShapes()) {
            handleTextShape(shape, data);
        }
    }

    /**
     * Xử lý thay thế image placeholders trong slide
     */
    private void replaceImagesInSlide(XSLFSlide slide, Map<String, MultipartFile> images, XMLSlideShow ppt) throws IOException {
        if (images == null || images.isEmpty()) {
            return;
        }

        List<XSLFShape> shapesToRemove = new ArrayList<>();
        List<ImagePlaceholder> imagesToAdd = new ArrayList<>();

        // Tìm tất cả text shapes có chứa image placeholders
        for (XSLFShape shape : slide.getShapes()) {
            findImagePlaceholders(shape, images, shapesToRemove, imagesToAdd);
        }

        // Thêm images mới vào slide
        for (ImagePlaceholder imgPlaceholder : imagesToAdd) {
            addImageToSlide(slide, ppt, imgPlaceholder);
        }

        // Xóa các text shapes đã được thay thế
        for (XSLFShape shape : shapesToRemove) {
            slide.removeShape(shape);
        }
    }

    /**
     * Tìm image placeholders trong shapes
     */
    private void findImagePlaceholders(XSLFShape shape, Map<String, MultipartFile> images,
                                     List<XSLFShape> shapesToRemove, List<ImagePlaceholder> imagesToAdd) {
        if (shape instanceof XSLFTextShape textShape) {
            String text = extractTextFromShape(textShape);

            // Kiểm tra xem có image placeholder không
            for (Map.Entry<String, MultipartFile> entry : images.entrySet()) {
                String key = entry.getKey();
                String placeholder = "{IMAGE:" + key + "}";

                if (text.contains(placeholder)) {
                    log.info("Found image placeholder '{}' in text: '{}'", placeholder, text);

                    // Lưu thông tin vị trí và kích thước của shape để thay thế
                    Rectangle2D anchor = shape.getAnchor();
                    ImagePlaceholder imgPlaceholder = new ImagePlaceholder(
                        key, entry.getValue(), anchor.getX(), anchor.getY(),
                        anchor.getWidth(), anchor.getHeight()
                    );
                    imagesToAdd.add(imgPlaceholder);
                    shapesToRemove.add(shape);
                    break; // Chỉ xử lý một placeholder mỗi shape
                }
            }
        }
        // Xử lý group shapes
        else if (shape instanceof XSLFGroupShape group) {
            for (XSLFShape inner : group.getShapes()) {
                findImagePlaceholders(inner, images, shapesToRemove, imagesToAdd);
            }
        }
        // Xử lý table cells
        else if (shape instanceof XSLFTable table) {
            for (XSLFTableRow row : table.getRows()) {
                for (XSLFTableCell cell : row.getCells()) {
                    String cellText = extractTextFromCell(cell);

                    for (Map.Entry<String, MultipartFile> entry : images.entrySet()) {
                        String key = entry.getKey();
                        String placeholder = "{IMAGE:" + key + "}";

                        if (cellText.contains(placeholder)) {
                            log.info("Found image placeholder '{}' in table cell", placeholder);

                            // Xóa text trong cell và thêm image
                            clearCellText(cell);
                            Rectangle2D cellAnchor = cell.getAnchor();
                            ImagePlaceholder imgPlaceholder = new ImagePlaceholder(
                                key, entry.getValue(), cellAnchor.getX(), cellAnchor.getY(),
                                cellAnchor.getWidth(), cellAnchor.getHeight()
                            );
                            imagesToAdd.add(imgPlaceholder);
                            break;
                        }
                    }
                }
            }
        }
    }

    /**
     * Thêm image vào slide tại vị trí đã chỉ định
     */
//    private void addImageToSlide(XSLFSlide slide, XMLSlideShow ppt, ImagePlaceholder imgPlaceholder) throws IOException {
//        try {
//            MultipartFile imageFile = imgPlaceholder.getFile();
//            byte[] imageData = imageFile.getBytes();
//
//            // Xác định định dạng ảnh
//            PictureData.PictureType pictureType = determinePictureType(imageFile.getOriginalFilename());
//
//            // Thêm ảnh vào presentation
//            XSLFPictureData pictureData = ppt.addPicture(imageData, pictureType);
//            XSLFPictureShape pictureShape = slide.createPicture(pictureData);
//
//            // Đặt vị trí và kích thước cho ảnh
//            Rectangle2D anchor = new Rectangle2D.Double(
//                imgPlaceholder.getX(),
//                imgPlaceholder.getY(),
//                imgPlaceholder.getWidth(),
//                imgPlaceholder.getHeight()
//            );
//            pictureShape.setAnchor(anchor);
//
//            log.info("Successfully added image '{}' to slide at position ({}, {}) with size {}x{}",
//                    imageFile.getOriginalFilename(),
//                    imgPlaceholder.getX(), imgPlaceholder.getY(),
//                    imgPlaceholder.getWidth(), imgPlaceholder.getHeight());
//
//        } catch (Exception e) {
//            log.error("Failed to add image for placeholder '{}': {}", imgPlaceholder.getKey(), e.getMessage());
//        }
//    }
    private void addImageToSlide(XSLFSlide slide, XMLSlideShow ppt, ImagePlaceholder imgPlaceholder) throws IOException {
        try {
            MultipartFile imageFile = imgPlaceholder.getFile();
            byte[] imageData = imageFile.getBytes();

            // Đọc ảnh gốc để lấy kích thước
            BufferedImage bimg = ImageIO.read(imageFile.getInputStream());
            if (bimg == null) {
                log.error("Invalid image format for file: {}", imageFile.getOriginalFilename());
                return;
            }
            int imgWidth = bimg.getWidth();
            int imgHeight = bimg.getHeight();

            // Xác định định dạng ảnh
            PictureData.PictureType pictureType = determinePictureType(imageFile.getOriginalFilename());

            // Thêm ảnh vào presentation
            XSLFPictureData pictureData = ppt.addPicture(imageData, pictureType);
            XSLFPictureShape pictureShape = slide.createPicture(pictureData);

            // Lấy khung anchor từ placeholder
            double anchorX = imgPlaceholder.getX();
            double anchorY = imgPlaceholder.getY();
            double anchorWidth = imgPlaceholder.getWidth();
            double anchorHeight = imgPlaceholder.getHeight();

            // Tính tỉ lệ scale để fit vào khung
            double scaleX = anchorWidth / imgWidth;
            double scaleY = anchorHeight / imgHeight;
            double scale = Math.min(scaleX, scaleY); // scale nhỏ nhất để không vượt khung

            // Kích thước mới sau khi scale
            double newWidth = imgWidth * scale;
            double newHeight = imgHeight * scale;

            // Căn giữa trong khung
            double newX = anchorX + (anchorWidth - newWidth) / 2;
            double newY = anchorY + (anchorHeight - newHeight) / 2;

            // Đặt vị trí + size
            Rectangle2D newAnchor = new Rectangle2D.Double(newX, newY, newWidth, newHeight);
            pictureShape.setAnchor(newAnchor);

            log.info("Successfully added image '{}' ({}x{}) scaled to {}x{} at position ({}, {})",
                    imageFile.getOriginalFilename(), imgWidth, imgHeight,
                    newWidth, newHeight, newX, newY);

        } catch (Exception e) {
            log.error("Failed to add image for placeholder '{}': {}", imgPlaceholder.getKey(), e.getMessage(), e);
        }
    }


    /**
     * Xác định loại ảnh dựa trên tên file
     */
    private PictureData.PictureType determinePictureType(String filename) {
        if (filename == null) {
            return PictureData.PictureType.PNG; // Default
        }

        String extension = filename.substring(filename.lastIndexOf('.') + 1).toLowerCase();
        return switch (extension) {
            case "jpg", "jpeg" -> PictureData.PictureType.JPEG;
            case "png" -> PictureData.PictureType.PNG;
            case "gif" -> PictureData.PictureType.GIF;
            case "bmp" -> PictureData.PictureType.BMP;
            case "tiff", "tif" -> PictureData.PictureType.TIFF;
            default -> PictureData.PictureType.PNG;
        };
    }

    /**
     * Trích xuất text từ một text shape
     */
    private String extractTextFromShape(XSLFTextShape textShape) {
        StringBuilder text = new StringBuilder();
        for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                String runText = run.getRawText();
                if (runText != null) {
                    text.append(runText);
                }
            }
        }
        return text.toString();
    }

    /**
     * Trích xuất text từ table cell
     */
    private String extractTextFromCell(XSLFTableCell cell) {
        StringBuilder text = new StringBuilder();
        for (XSLFTextParagraph paragraph : cell.getTextParagraphs()) {
            for (XSLFTextRun run : paragraph.getTextRuns()) {
                String runText = run.getRawText();
                if (runText != null) {
                    text.append(runText);
                }
            }
        }
        return text.toString();
    }

    /**
     * Xóa text trong table cell
     */
    private void clearCellText(XSLFTableCell cell) {
        cell.clearText();
    }

    /**
     * Đệ quy xử lý shape (text, group, bảng) với cải thiện xử lý text runs
     */
    private void handleTextShape(XSLFShape shape, Map<String, String> data) {
        if (shape instanceof XSLFTextShape textShape) {
            // Cải thiện: xử lý từng paragraph một cách toàn diện
            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                replacePlaceholdersInParagraph(paragraph, data);
            }
        }
        // Nếu shape là group thì duyệt tiếp
        else if (shape instanceof XSLFGroupShape group) {
            for (XSLFShape inner : group.getShapes()) {
                handleTextShape(inner, data);
            }
        }
        // Nếu shape là bảng
        else if (shape instanceof XSLFTable table) {
            for (XSLFTableRow row : table.getRows()) {
                for (XSLFTableCell cell : row.getCells()) {
                    for (XSLFTextParagraph paragraph : cell.getTextParagraphs()) {
                        replacePlaceholdersInParagraph(paragraph, data);
                    }
                }
            }
        }
    }

    /**
     * Xử lý placeholder trong paragraph một cách thông minh hơn
     * Phương pháp này sẽ gộp toàn bộ text của paragraph, thay thế, rồi gán lại
     */
    private void replacePlaceholdersInParagraph(XSLFTextParagraph paragraph, Map<String, String> data) {
        List<XSLFTextRun> runs = paragraph.getTextRuns();
        if (runs.isEmpty()) {
            return;
        }

        // Gộp toàn bộ text từ tất cả runs
        StringBuilder fullText = new StringBuilder();
        for (XSLFTextRun run : runs) {
            String text = run.getRawText();
            if (text != null) {
                fullText.append(text);
            }
        }

        String originalText = fullText.toString();
        String replacedText = originalText;

        // Debug: in ra text gốc
        if (!originalText.trim().isEmpty()) {
            log.debug("Original text: " + originalText);
        }

        // Thay thế tất cả placeholders
        for (Map.Entry<String, String> entry : data.entrySet()) {
            String placeholder = "{" + entry.getKey() + "}";
            replacedText = replacedText.replace(placeholder, entry.getValue());
        }

        // Nếu có thay đổi, cập nhật lại paragraph
        if (!originalText.equals(replacedText)) {
            log.debug("Replaced text: " + replacedText);

            // Xóa tất cả runs hiện tại trừ run đầu tiên
            XSLFTextRun firstRun = runs.get(0);

            // Xóa các runs khác
            for (int i = runs.size() - 1; i > 0; i--) {
                // Không thể xóa trực tiếp, nên ta sẽ set text rỗng
                runs.get(i).setText("");
            }

            // Gán text mới vào run đầu tiên
            firstRun.setText(replacedText);
        }
    }

    /**
     * Class để lưu thông tin image placeholder
     */
    @Getter
    private static class ImagePlaceholder {
        private final String key;
        private final MultipartFile file;
        private final double x, y, width, height;

        public ImagePlaceholder(String key, MultipartFile file, double x, double y, double width, double height) {
            this.key = key;
            this.file = file;
            this.x = x;
            this.y = y;
            this.width = width;
            this.height = height;
        }

    }
}
