package bean2Excel;

import bean2Excel.style.GeneralCellStyle;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import java.util.Random;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public class Bean2ExcelTest {

    @Getter
    @AllArgsConstructor
    @GeneralCellStyle
    public static class ClassA {
        @ExcelColumn(columnName = "column 1", cellType = CellType.STRING, order = 100)
        private String column1;
        @ExcelColumn(columnName = "column 2", cellType = CellType.NUMERIC, order = 5)
        private double column2;
        @ExcelColumn(columnName = "column 3", cellType = CellType.BOOLEAN, order = 3)
        private boolean column3;
        @ExcelColumn(columnName = "column 4", cellType = CellType.NUMERIC, order = 1,
            valueConverter = String2DoubleConverter.class)
        private String column4;

    }

    public static class String2DoubleConverter implements ValueConverter {
        @Override
        public Object convert(Object o) {
            return Double.parseDouble((String) o);
        }
    }

    static int TEST_DATA_COUNT = 1000;
    static ClassA[] testData;
    @BeforeAll
    static void generateData() {
        testData = new ClassA[TEST_DATA_COUNT];
        Random random = new Random();

        for (int i = 0; i < TEST_DATA_COUNT; i++) {
            testData[i] = new ClassA(
                getRandomString(100, random),
                random.nextDouble() * 1000,
                random.nextBoolean(),
                Double.toString(random.nextDouble()*1000)
            );
        }
    }

    @Test
    void happyDays() throws Exception {

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = Bean2Excel.getCreateSheetFunc(ClassA.class).exec(List.of(testData), workbook, "My sheet");

            Row header = sheet.getRow(0);
            assertEquals("column 4", header.getCell(0).getStringCellValue());
            assertEquals("column 3", header.getCell(1).getStringCellValue());
            assertEquals("column 2", header.getCell(2).getStringCellValue());
            assertEquals("column 1", header.getCell(3).getStringCellValue());

            for (int i = 0; i < TEST_DATA_COUNT; i++) {
                Row curRow = sheet.getRow(i+1);
                ClassA curObject = testData[i];
                assertEquals(curObject.getColumn1(), curRow.getCell(3).getStringCellValue());
                assertEquals(curObject.getColumn2(), curRow.getCell(2).getNumericCellValue(), 0.1);
                assertEquals(curObject.isColumn3(), curRow.getCell(1).getBooleanCellValue());
                assertEquals(curObject.getColumn4(), Double.toString(curRow.getCell(0).getNumericCellValue()));
            }
        }
    }

    private static String getRandomString(int length, Random randomGen) {
        StringBuilder out = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            out.append((char) randomGen.nextInt(65535));
        }
        return out.toString();
    }

    @Test
    void testSpeed() throws Exception {

        Bean2Excel.clearCache();

        try (Workbook workbook = new XSSFWorkbook()) {

            long nonCacheSpeed = TestUtil.measureSpeed("create sheet non cached", ()->{
                Bean2Excel.getCreateSheetFunc(ClassA.class).exec(List.of(testData), workbook, "My sheet");
            });
            workbook.removeSheetAt(0);
            long cachedSpeed = TestUtil.measureSpeed("create sheet cached", ()->{
                Bean2Excel.getCreateSheetFunc(ClassA.class).exec(List.of(testData), workbook, "My sheet");
            });
            workbook.removeSheetAt(0);

            TestUtil.measureSpeed("create sheet cached (another)", ()->{
                Bean2Excel.getCreateSheetFunc(ClassA.class).exec(List.of(testData), workbook, "My sheet");
            });

            assertTrue(nonCacheSpeed > cachedSpeed);
        }
    }
}
