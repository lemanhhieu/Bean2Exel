# Bean2Excel

## Usage

### Export list of java beans to an Excel sheet

```
@Data // <-- used lombok 
public class ClassA {
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

public class String2DoubleConverter implements ValueConverter {
    @Override
    public Object convert(Object o) {
        return Double.parseDouble((String) o);
    }
}

...
List<ClassA> objectList = ... // initialize a list
Workbook workbook = new XSSFWorkbook(); // or whatever kind of workbook you prefer
Sheet sheet = Bean2Excel.getCreateSheetFunc(ClassA.class).exec(objectList, workbook, "Put your sheet name here");
```

Above code will produce an Excel sheet with columns' names and cell types as annotated in your Java Beans.
The order of the columns goes from smallest (left) to largest (right).

### Customize cell style

To customize cell style and header cell style of each column, 
use `CellStyleProperties` annotation for column cell style and `GeneralCellStyle` annotation for header cell style.
Those annotations require you to create provider classes that **must have a no args constructor**.
Do not modify the cells inside your provider classes.

Example usage:

```
@Data // <-- used lombok
@GeneralCellStyle(cellStyle = MyCellStyleProvider.class)
public class ClassA {
    @ExcelColumn(...)
    @CellStyleProperties(headerCellStyle = MyCellStylePropertiesProvider.class, autofit = false)
    private String column1;
    
    @ExcelColumn(...)
    @CellStyleProperties(cellStyle = MyCellStylePropertiesProvider.class)
    private double column2;
}

public class MyCellStyleProvider implements CellStyleProvider{
    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        return cellStyle;
    }
}

public class MyCellStylePropertiesProvider implements CellStylePropertiesProvider {
    @Override
    public Map<String, Object> getCellStyleProperties(
        Cell cell // you can use this to customize your style based on cell value
    ) {
        Map<String, Object> properties = new HashMap<String, Object>();
        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
        return properties;
    }
}

```
See Apache POI's documentation for info on cell style and cell style properties:
* [Cell style properties](https://poi.apache.org/components/spreadsheet/quick-guide.html#CellProperties)
* [Cell style](https://poi.apache.org/components/spreadsheet/quick-guide.html#Borders)