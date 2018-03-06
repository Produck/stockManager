package util;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public interface RowReader {
    public void processing(Row row, Sheet sheet);
}
