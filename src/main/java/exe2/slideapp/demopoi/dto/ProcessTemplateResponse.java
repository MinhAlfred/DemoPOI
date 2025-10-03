package exe2.slideapp.demopoi.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ProcessTemplateResponse {
    private String message;
    private String filename;
    private long fileSize;
    private int textReplacementsCount;
    private int imageReplacementsCount;
    private LocalDateTime processedAt;
    private boolean success;
}