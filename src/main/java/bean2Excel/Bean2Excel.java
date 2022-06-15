/*
 * MIT License
 *
 * Copyright (c) 2022 Le Manh Hieu
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

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
        /**
         * Create an Excel sheet from a list of Java Beans objects. Each object is the data of a row.
         * @param objectList a list of Java Beans objects
         * @param workbook the workbook used to create Excel sheet
         * @param sheetName name of the sheet to be created
         * @return the created sheet
         * @throws Bean2ExcelException a runtime exception indicating that a declaration or usage is invalid.
         */
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

    private static final Map<Class<?>, CreateSheetFunc<?>> cache = new Hashtable<>();

    /**
     * Clear cache. Intended to be used only for testing purpose
     */
    public static void clearCache() {
        cache.clear();
    }

    /**
     * Use to get a function to create Excel sheet.
     * <br/>
     * See {@link CreateSheetFunc#exec(List, Workbook, String)}.
     * @param objectType Type of object to be used as java beans
     * @return A function to create Excel sheet
     * @throws Bean2ExcelException a runtime exception indicating that a declaration or usage is invalid.
     */
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
