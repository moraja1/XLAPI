package api.xl.base;

import api.xl.constants.XLTags;
import api.xl.exceptions.XLFactoryException;

class XLFactoryTest {
    public static void main(String[] args) {
        try {
            XLWorkbook w = XLFactory.newWorkbook();
        } catch (XLFactoryException e) {
            throw new RuntimeException(e);
        }

    }
}