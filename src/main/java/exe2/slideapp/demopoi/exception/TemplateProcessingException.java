package exe2.slideapp.demopoi.exception;

import lombok.Getter;

import java.util.ArrayList;
import java.util.List;

@Getter
class TemplateProcessingException extends RuntimeException {
    private final List<String> details;

    public TemplateProcessingException(String message) {
        super(message);
        this.details = new ArrayList<>();
    }

    public TemplateProcessingException(String message, Throwable cause) {
        super(message, cause);
        this.details = new ArrayList<>();
        if (cause != null) {
            this.details.add("Cause: " + cause.getMessage());
        }
    }

    public TemplateProcessingException(String message, List<String> details) {
        super(message);
        this.details = details != null ? details : new ArrayList<>();
    }

    public void addDetail(String detail) {
        this.details.add(detail);
    }
}
