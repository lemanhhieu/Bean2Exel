package bean2Excel.style;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Map;

public interface CellStylePropertiesProvider {

    Map<String, Object> getCellStyleProperties(Cell cell);
}
