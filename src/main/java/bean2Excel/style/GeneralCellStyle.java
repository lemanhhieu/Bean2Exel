package bean2Excel.style;

import bean2Excel.style.CellStyleProvider;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface GeneralCellStyle {
    Class<? extends CellStyleProvider> cellStyle() default DefaultCellStyle.class;
}
