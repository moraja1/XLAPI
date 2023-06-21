package una.filesorganizeridoffice.business.api.xl.annotations;

import java.lang.annotation.Repeatable;

@Repeatable(XLCellSetValue.class)
public @interface XLCellColumn {
    int processOf();
    String column();
}
