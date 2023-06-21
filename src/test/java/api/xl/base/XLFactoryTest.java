package api.xl.base;

import api.xl.XLWorkbook;
import api.xl.exceptions.XLFactoryException;

import java.io.File;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.net.URL;

import static org.junit.jupiter.api.Assertions.*;

class XLFactoryTest {
    public static void main(String[] args) {
        URL url = XLWorkbook.class.getResource("EmptyExcel.xlsx");
        XLWorkbook workbook;
        try {
            File xlsx = new File(url.toURI());
            workbook = XLFactory.buildWorkbook(xlsx);
        } catch (URISyntaxException e) {
            throw new RuntimeException(e);
        } catch (XLFactoryException e) {
            throw new RuntimeException(e);
        }
    }
}