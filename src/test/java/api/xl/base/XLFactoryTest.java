package api.xl.base;

import api.xl.exceptions.XLFactoryException;

import java.io.File;
import java.net.URISyntaxException;
import java.net.URL;

class XLFactoryTest {
    public static void main(String[] args) {
        URL url = XLWorkbook.class.getResource("../Book1.xlsx");
        XLWorkbook workbook;
        try {
            File xlsx = new File(url.toURI());
            workbook = XLFactory.openWorkbook(xlsx);
        } catch (URISyntaxException e) {
            throw new RuntimeException(e);
        } catch (XLFactoryException e) {
            throw new RuntimeException(e);
        }
    }
}