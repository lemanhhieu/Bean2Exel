package bean2Excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.CellType;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.util.Optional;

import static bean2Excel.BeanInfo.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

@TestInstance(value = TestInstance.Lifecycle.PER_CLASS)
public class BeanInfoTest {


    @NoArgsConstructor
    @AllArgsConstructor
    @Getter
    public static class ClassA {
        @ExcelColumn(columnName = "column 1", cellType = CellType.STRING)
        private String column1;
        @ExcelColumn(columnName = "column 2", cellType = CellType.NUMERIC)
        private Double column2;
        @ExcelColumn(columnName = "column 3", cellType = CellType.BOOLEAN)
        private boolean column3;
    }

    @Test
    void happyDays() {
        ClassA objectA = new ClassA("foo", 2.3, true);
        ExcelObjectInfo excelObjectInfo = getExcelInfoFromBeans(objectA.getClass());
        Optional<FieldInfo> firstFieldInfo = excelObjectInfo.fieldInfoList().stream()
            .filter((fieldInfo)->
                fieldInfo.columnInfo().columnName().equals("column 1")
                && fieldInfo.columnInfo().cellType().equals(CellType.STRING)
            )
            .findAny();
        assertTrue(firstFieldInfo.isPresent());
        assertEquals("foo", firstFieldInfo.get().getter().exec(objectA));

        Optional<FieldInfo> secondFieldInfo = excelObjectInfo.fieldInfoList().stream()
            .filter((fieldInfo)->
                fieldInfo.columnInfo().columnName().equals("column 2")
                    && fieldInfo.columnInfo().cellType().equals(CellType.NUMERIC)
            )
            .findAny();
        assertTrue(secondFieldInfo.isPresent());
        assertEquals(2.3, secondFieldInfo.get().getter().exec(objectA));

        Optional<FieldInfo> thirdFieldInfo = excelObjectInfo.fieldInfoList().stream()
            .filter((fieldInfo)->
                fieldInfo.columnInfo().columnName().equals("column 3")
                    && fieldInfo.columnInfo().cellType().equals(CellType.BOOLEAN)
            )
            .findAny();
        assertTrue(thirdFieldInfo.isPresent());
        assertEquals(true, thirdFieldInfo.get().getter().exec(objectA));
    }
}
