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

import bean2Excel.style.CellStyleProperties;
import bean2Excel.style.GeneralCellStyle;
import lombok.NonNull;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

class BeanInfo {

    record ExcelObjectInfo(
        @Nullable GeneralCellStyle generalStyle,
        @NotNull List<FieldInfo> fieldInfoList
    ) {}

    record FieldInfo(
        @NotNull GetterFunc<?> getter,
        @NotNull ExcelColumn columnInfo,
        @Nullable CellStyleProperties cellStyleProperties
    ) {}

    @FunctionalInterface
    interface GetterFunc<T> {
        T exec(Object o);
    }

    public static ExcelObjectInfo getExcelInfoFromBeans(@NonNull Class<?> clazz) {

        ArrayList<FieldInfo> fieldInfoList = new ArrayList<>();

        for (Class<?> curClass = clazz; curClass != null; curClass = curClass.getSuperclass()) {
            for (Field field : curClass.getDeclaredFields()) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                if (excelColumn != null) {
                    Method getterMethod = getGetter(field);
                    GetterFunc<?> getterFunc = (o)->{
                        try {
                            return getterMethod.invoke(o);
                        }
                        catch (Exception e) {
                            e.printStackTrace();
                            throw new Bean2ExcelException("Internal Error");
                        }
                    };
                    fieldInfoList.add(new FieldInfo(getterFunc, excelColumn, field.getAnnotation(CellStyleProperties.class)));
                }
            }
        }

        // sort field based on order attribute of the annotation
        fieldInfoList.sort(new ColumnComparator());

        return new ExcelObjectInfo(
            clazz.getAnnotation(GeneralCellStyle.class),
            fieldInfoList
        );
    }


    private static Method getGetter(Field field) {
        String verb = field.getType().isPrimitive() && field.getType().equals(boolean.class) ?
            "is" : "get";
        String getterName = verb +
            Character.toUpperCase(field.getName().charAt(0)) +
            field.getName().substring(1);
        try {
            return field.getDeclaringClass().getMethod(getterName);
        }
        catch (NoSuchMethodException e) {
            throw new Bean2ExcelException(String.format(
                "Can't find public getter \"%s\" of \"%s\" for field \"%s\"",
                getterName, field.getDeclaringClass().getName(), field.getName()));
        }
    }

    static class ColumnComparator implements Comparator<FieldInfo> {

        @Override
        public int compare(FieldInfo fieldInfo1, FieldInfo fieldInfo2) {
            return Integer.compare(fieldInfo1.columnInfo().order(), fieldInfo2.columnInfo().order());
        }
    }
}
