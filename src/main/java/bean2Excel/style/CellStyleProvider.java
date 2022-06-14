package bean2Excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellStyleProvider {
    CellStyle getCellStyle(Workbook workbook);
}
