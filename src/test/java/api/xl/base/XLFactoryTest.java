package api.xl.base;

import api.xl.exceptions.XLFactoryException;

import java.io.File;
import java.io.IOException;

class XLFactoryTest {
    public static void main(String[] args) {
        try {
            XLWorkbook w = XLFactory.openWorkbook(new File("C:\\Users\\N00148095\\Desktop\\formatAdult.xlsx"));
            XLFactory.saveWorkbook(w, "C:\\Users\\N00148095\\Desktop\\formatAdultCopy.xlsx");
        } catch (XLFactoryException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}