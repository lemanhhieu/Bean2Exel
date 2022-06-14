package bean2Excel;

import bean2Excel.style.CellStylePropertiesProvider;
import lombok.NonNull;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.util.*;
import java.util.function.Consumer;

import static bean2Excel.BeanInfo.*;

public class Bean2Excel {
    @FunctionalInterface
    public interface CreateSheetFunc<T> {
        Sheet exec(
            @NonNull List<T> objectList,
            @NonNull Workbook workbook,
            @NonNull String sheetName
        );
    }
    private record ProcessedFieldInfo(
        @NotNull FieldInfo fieldInfo,
        @NotNull Integer columnIndex,
        @NotNull ValueConverter valueConverter,
        @Nullable StylePropertiesSetter styleSetter

    ){}
    private record StylePropertiesSetter(
        @NotNull Consumer<Cell> headerStylePropertiesSetter,
        @NotNull Consumer<Cell> cellStylePropertiesSetter
    ) {}

    private record StyleProviders(
        List<? extends CellStylePropertiesProvider> headerStyleProviders,
        List<? extends CellStylePropertiesProvider> cellStyleProviders
    ){}

    private static final Map<Class<?>, CreateSheetFunc<?>> cache = new Hashtable<>();
    public static void clearCache() {
        cache.clear();
    }
    public static <T> CreateSheetFunc<T> getCreateSheetFunc(Class<T> objectType) {
        // try to use cache
        val cachedResult = cache.get(objectType);
        if (cachedResult != null) {
            return (CreateSheetFunc<T>) cachedResult;
        }

        ExcelObjectInfo excelObjectInfo = getExcelInfoFromBeans(objectType);
        Map<String, ProcessedFieldInfo> processedInfo = processFieldsInfo(excelObjectInfo);

        CreateSheetFunc<T> createSheetFunc = (objectList, workbook, sheetName) ->{
            Sheet sheet = workbook.createSheet(sheetName);
            CellStyle generalStyle = null;
            if (excelObjectInfo.generalStyle() != null){
                generalStyle = getNoArgsInstance(excelObjectInfo.generalStyle().cellStyle()).getCellStyle(workbook);
            }

            // create header row
            Row headerRow = sheet.createRow(0);
            for (val processed : processedInfo.entrySet()) {
                Cell headerCell = headerRow.createCell(processed.getValue().columnIndex());
                headerCell.setCellStyle(generalStyle);
                if (processed.getValue().styleSetter() != null) {
                    processed.getValue().styleSetter().headerStylePropertiesSetter().accept(headerCell);
                }
                headerCell.setCellValue(processed.getKey());
            }

            // create data rows
            int rowIndex = 1;
            for (val rowObject : objectList) {
                Row curRow = sheet.createRow(rowIndex);

                for (val processed : processedInfo.entrySet()) {
                    Cell cell = curRow.createCell(processed.getValue().columnIndex());

                    // set style
                    cell.setCellStyle(generalStyle);
                    if (processed.getValue().styleSetter() != null) {
                        processed.getValue().styleSetter().cellStylePropertiesSetter().accept(cell);
                    }

                    // set value
                    val cellValue = processed.getValue()
                        .valueConverter()
                        .convert(processed.getValue().fieldInfo().getter().exec(rowObject));

                    if (cellValue == null) {
                        cell.setBlank();
                    } else {
                        switch (processed.getValue().fieldInfo().columnInfo().cellType()) {
                            case STRING -> cell.setCellValue((String) cellValue);
                            case BOOLEAN -> cell.setCellValue((Boolean) cellValue);
                            case NUMERIC -> cell.setCellValue((Double) cellValue);
                            case BLANK -> cell.setBlank();
                            default -> throw new Bean2ExcelException(
                                String.format("Unsupported excel type \"%s\"",
                                    processed.getValue().fieldInfo().columnInfo().cellType())
                            );
                        }
                    }
                }
                rowIndex++;
            }

            // autofit column
            for (val processed : processedInfo.entrySet()) {
                if (processed.getValue().fieldInfo().cellStyleProperties() != null
                    && processed.getValue().fieldInfo().cellStyleProperties().autoFit()
                ) {
                    sheet.autoSizeColumn(processed.getValue().columnIndex());
                }
            }

            return sheet;
        };

        cache.put(objectType, createSheetFunc);
        return createSheetFunc;
    }

    private static Map<String, ProcessedFieldInfo> processFieldsInfo(ExcelObjectInfo excelObjectInfo) {
        Map<String, ProcessedFieldInfo> processedInfo = new HashMap<>();

        int columnIndex = 0;
        for (FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {

            if (processedInfo.containsKey(fieldInfo.columnInfo().columnName())) {
                throw new Bean2ExcelException(
                    String.format("Duplicate column name \"%s\"",
                        fieldInfo.columnInfo().columnName())
                );
            }

            StylePropertiesSetter styleSetter = null;

            if (fieldInfo.cellStyleProperties() != null) {
                val headerStyleProviders =
                    Arrays
                        .stream(fieldInfo.cellStyleProperties().headerCellStyle())
                        .map(Bean2Excel::getNoArgsInstance)
                        .toList();
                val cellStyleProviders =
                    Arrays.stream(fieldInfo.cellStyleProperties().cellStyle())
                        .map(Bean2Excel::getNoArgsInstance)
                        .toList();

                styleSetter = new StylePropertiesSetter(
                    (cell) -> CellUtil.setCellStyleProperties(cell, mergePropertiesMap(headerStyleProviders, cell)),
                    (cell) -> CellUtil.setCellStyleProperties(cell, mergePropertiesMap(cellStyleProviders, cell))
                );
            }
            val valueConverter = getNoArgsInstance(fieldInfo.columnInfo().valueConverter());

            processedInfo.put(
                fieldInfo.columnInfo().columnName(),
                new ProcessedFieldInfo(fieldInfo, columnIndex, valueConverter, styleSetter)
            );
            columnIndex++;
        }

        return processedInfo;
    }

    private static Map<String, Object> mergePropertiesMap(
        List<? extends CellStylePropertiesProvider> providerList,
        Cell cell
    ) {
        Map<String, Object> mergedProperties = new HashMap<>();
        for (val provider : providerList) {
            mergedProperties.putAll(provider.getCellStyleProperties(cell));
        }
        return mergedProperties;
    }

    private static <T> T getNoArgsInstance(Class<T> clazz) {
        try {
            return clazz.getConstructor().newInstance();
        }
        catch (ReflectiveOperationException e) {
            throw new Bean2ExcelException(String.format(
                "Failed to use instantiate class \"%s\" using no args constructor, either because it doesn't exists, or the class is not public",
                clazz.getName()
            ));
        }
    }

}
