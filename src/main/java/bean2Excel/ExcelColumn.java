package bean2Excel;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {
    String columnName();
    CellType cellType();
    int order() default 0;
    Class<? extends ValueConverter> valueConverter() default IdentityValueConverter.class;
}
