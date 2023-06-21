package api.xl.util;

import java.time.LocalDate;

public final class Converter {
    public static String cellTypetoString(Object cellValue) {
        if(cellValue instanceof Integer){
            return cellValue.toString();
        }
        else if(cellValue instanceof LocalDate){
            return DateUtil.toString(((LocalDate)cellValue));
        }
        else if(cellValue instanceof String){
            return (String)cellValue;
        }
        else if(cellValue instanceof Float){
            return String.valueOf(cellValue);
        }
        else{
            return "";
        }
    }
}
