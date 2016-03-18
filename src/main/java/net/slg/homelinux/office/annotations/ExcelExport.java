/**
 * @author VZMUELLC
 */

package net.slg.homelinux.office.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;
import java.lang.annotation.ElementType;
import java.lang.annotation.RetentionPolicy;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.METHOD)
public @interface ExcelExport {

	String columnName() default "";
	String dateFormat() default "dd.MM.yyyy";
	String suppressedBy() default "";
	String subItem() default "";
	int order() default -1;
	
}
