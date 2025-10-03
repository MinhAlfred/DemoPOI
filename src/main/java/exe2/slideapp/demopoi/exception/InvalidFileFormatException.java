package exe2.slideapp.demopoi.exception;

import lombok.Getter;
import java.util.ArrayList;
import java.util.List;

/**
 * Exception thrown when file format is invalid
 */
@Getter
public class InvalidFileFormatException extends RuntimeException {
    private final String fileName;

    public InvalidFileFormatException(String message, String fileName) {
        super(message);
        this.fileName = fileName;
    }

    public InvalidFileFormatException(String message) {
        super(message);
        this.fileName = null;
    }
}

