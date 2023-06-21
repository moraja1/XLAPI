package una.filesorganizeridoffice.business.api.xl.util;

import java.math.BigDecimal;

public final class NumberUtil {
    /***
     * Detects whether a String contains a Scientific Notation Number
     * @param numberString String vale
     * @return Boolean
     */
    public static boolean isScientificNotation(String numberString) {
        // Validate number
        try {
            new BigDecimal(numberString);
        } catch (NumberFormatException e) {
            return false;
        }
        // Check for scientific notation
        return numberString.toUpperCase().contains("E") && (numberString.charAt(1)=='.' || numberString.charAt(2)=='.');
    }

    public static boolean isNumber(String value){
        return value.matches("-?\\d+(\\.\\d+)?");
    }
}
