package exe2.slideapp.demopoi.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

@Service
@Slf4j
public class PowerPointService {

    public java.io.File processTemplate(MultipartFile file, Map<String, String> data) throws Exception {
        java.io.File tempFile = java.io.File.createTempFile("pptx-template", ".pptx");

        try (InputStream fis = file.getInputStream();
             XMLSlideShow ppt = new XMLSlideShow(fis);
             FileOutputStream fos = new FileOutputStream(tempFile)) {

            // Debug: in ra số lượng slides
            log.info("Total slides: " + ppt.getSlides().size());

            int slideIndex = 0;
            for (XSLFSlide slide : ppt.getSlides()) {
                System.out.println("Processing slide " + (slideIndex + 1));
                replaceTextInSlide(slide, data);
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
            handleShape(shape, data);
        }
    }

    /**
     * Đệ quy xử lý shape (text, group, bảng) với cải thiện xử lý text runs
     */
    private void handleShape(XSLFShape shape, Map<String, String> data) {
        if (shape instanceof XSLFTextShape textShape) {
            // Cải thiện: xử lý từng paragraph một cách toàn diện
            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                replacePlaceholdersInParagraph(paragraph, data);
            }
        }
        // Nếu shape là group thì duyệt tiếp
        else if (shape instanceof XSLFGroupShape group) {
            for (XSLFShape inner : group.getShapes()) {
                handleShape(inner, data);
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
            System.out.println("Original text: " + originalText);
        }

        // Thay thế tất cả placeholders
        for (Map.Entry<String, String> entry : data.entrySet()) {
            String placeholder = "{" + entry.getKey() + "}";
            replacedText = replacedText.replace(placeholder, entry.getValue());
        }

        // Nếu có thay đổi, cập nhật lại paragraph
        if (!originalText.equals(replacedText)) {
            System.out.println("Replaced text: " + replacedText);

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
}
