package bean2Excel.style;

import org.apache.poi.ss.usermodel.Cell;

import java.util.Collections;
import java.util.Map;

public class DefaultCellStyleProperties implements CellStylePropertiesProvider {
    @Override
    public Map<String, Object> getCellStyleProperties(Cell cell) {
        return Collections.emptyMap();
    }
}
