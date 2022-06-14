package bean2Excel.style;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface CellStyleProperties {
    Class<? extends CellStylePropertiesProvider>[] headerCellStyle() default { DefaultCellStyleProperties.class };
    Class<? extends CellStylePropertiesProvider>[] cellStyle() default { DefaultCellStyleProperties.class };
}
