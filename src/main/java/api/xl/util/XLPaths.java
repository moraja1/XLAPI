package api.xl.util;

import java.util.function.UnaryOperator;

public final class XLPaths {
    public static final String SHARED_STRINGS = "xl/sharedStrings.xml";
    public static final String STYLES = "xl/styles.xml";
    public static final String WORKBOOK = "xl/workbook.xml";
    public static final UnaryOperator<String> SHEET = i -> "xl/worksheets/sheet" + i + ".xml";
}
