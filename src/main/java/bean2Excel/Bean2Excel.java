package bean2Excel;

import bean2Excel.style.CellStylePropertiesProvider;
import lombok.NonNull;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.util.*;

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
    private record StyleProviders(
        List<? extends CellStylePropertiesProvider> headerStyleProviders,
        List<? extends CellStylePropertiesProvider> cellStyleProviders
    ){}

    private static final Map<Class<?>, CreateSheetFunc<?>> cache = new Hashtable<>();

    public static <T> CreateSheetFunc<T> getCreateSheetFunc(Class<T> objectType) {
        // try to use cache
        val cachedResult = cache.get(objectType);
        if (cachedResult != null) {
            return (CreateSheetFunc<T>) cachedResult;
        }

        ExcelObjectInfo excelObjectInfo = getExcelInfoFromBeans(objectType);

        // build style providers map for each column
        Map<String, StyleProviders> styleProvidersList = new HashMap<>();

        for (BeanInfo.FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {
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
                styleProvidersList.put(
                    fieldInfo.columnInfo().columnName(),
                    new StyleProviders(headerStyleProviders, cellStyleProviders)
                );
            }
        }

        // build column index map
        int columnIndex = 0;
        Map<String, Integer> columnIndexMap = new HashMap<>();
        for (FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {

            if (columnIndexMap.containsKey(fieldInfo.columnInfo().columnName())) {
                throw new Bean2ExcelException(
                    String.format("Duplicate column name \"%s\"",
                        fieldInfo.columnInfo().columnName())
                );
            }
            columnIndexMap.put(fieldInfo.columnInfo().columnName(), columnIndex);
            columnIndex++;
        }

        CreateSheetFunc<T> createSheetFunc = (objectList, workbook, sheetName) -> {
            Sheet sheet = workbook.createSheet(sheetName);
            CellStyle generalStyle = null;
            if (excelObjectInfo.generalStyle() != null){
                generalStyle = getNoArgsInstance(excelObjectInfo.generalStyle().cellStyle()).getCellStyle(workbook);
            }
            // create header row
            Row headerRow = sheet.createRow(0);
            for (BeanInfo.FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {
                Cell cell = headerRow.createCell(columnIndexMap.get(fieldInfo.columnInfo().columnName()));

                cell.setCellStyle(generalStyle);
                if (styleProvidersList.containsKey(fieldInfo.columnInfo().columnName())) {
                    val styleProviders = styleProvidersList
                        .get(fieldInfo.columnInfo().columnName())
                        .headerStyleProviders();
                    CellUtil.setCellStyleProperties(cell, mergePropertiesMap(styleProviders, cell));
                }

                cell.setCellValue(fieldInfo.columnInfo().columnName());
            }

            // create data rows
            int rowIndex = 1;
            for (val rowObject : objectList) {
                Row curRow = sheet.createRow(rowIndex);
                for (BeanInfo.FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {
                    Cell cell = curRow.createCell(columnIndexMap.get(fieldInfo.columnInfo().columnName()));

                    cell.setCellStyle(generalStyle);
                    if (styleProvidersList.containsKey(fieldInfo.columnInfo().columnName())) {
                        val styleProviders = styleProvidersList
                            .get(fieldInfo.columnInfo().columnName())
                            .cellStyleProviders();
                        CellUtil.setCellStyleProperties(cell, mergePropertiesMap(styleProviders, cell));
                    }


                    val valueConverter = getNoArgsInstance(fieldInfo.columnInfo().valueConverter());
                    val cellValue = valueConverter.convert(fieldInfo.getter().exec(rowObject));

                    if (cellValue == null) {
                        cell.setBlank();
                    }
                    else {
                        switch (fieldInfo.columnInfo().cellType()) {
                            case STRING -> cell.setCellValue((String) cellValue);
                            case BOOLEAN -> cell.setCellValue((Boolean) cellValue);
                            case NUMERIC -> cell.setCellValue((Double) cellValue);
                            case BLANK -> cell.setBlank();
                            default -> throw new Bean2ExcelException(
                                String.format("Unsupported excel type \"%s\"", fieldInfo.columnInfo().cellType())
                            );
                        }
                    }
                }
                rowIndex++;
            }

            for (BeanInfo.FieldInfo fieldInfo : excelObjectInfo.fieldInfoList()) {
                if (fieldInfo.columnInfo().autoFit()) {
                    sheet.autoSizeColumn(columnIndexMap.get(fieldInfo.columnInfo().columnName()));
                }
            }
            return sheet;
        };

        cache.put(objectType, createSheetFunc);
        return createSheetFunc;
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
