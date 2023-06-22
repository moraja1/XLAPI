package api.xl.base;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.zip.ZipFile;

import static api.xl.util.XLNumberUtil.isNumber;

/**
 * This class represents a xlsx workbook. This class contains main xlsx archives ass sheets, sharedString and styles and performs
 * changes into sharedString.xml file, such ass adding sharedString and returning sharedString index. It should also perform
 * other tasks such as managing styles.
 */
public abstract class XLWorkbook {
    private final String xlName;
    private final ZipFile xlFile;
    private final Document xlWorkbook;
    protected final Document xlSharedStrings;
    private final Document xlStyles;
    private final List<XLSheet> xlSheets = new LinkedList<>();
    /**
     * Creates a XLWorkbook and captures its name.
     * @param xlsx
     */
    public XLWorkbook(ZipFile xlsx, Document xlWorkbook, Document xlSharedStrings, Document xlStyles) {
        xlFile = xlsx;
        xlName = xlsx.getName();
        this.xlWorkbook = xlWorkbook;
        this.xlSharedStrings = xlSharedStrings;
        this.xlStyles = xlStyles;
    }
    /**
     * Returns the name of the XLWorkbook
     * @return String with the xlsx name.
     */
    public String getXlName() {
        return xlName;
    }
    /**
     * Return the Path object that represent the xlsx file.
     * @return Path of the xlsx file
     */
    public ZipFile getXlFile() {
        return xlFile;
    }
    /**
     * Return a Document that contains all the xl/workbook.xml information.
     * @return Document xl/workbook.xml
     */
    public Document getXlWorkbook() {
        return xlWorkbook;
    }

    /**
     * Return a Document the contains all the xl/sharedStrings.xml information.
     * @return Document xl/sharedStrings.xml
     */
    public Document getXlSharedStrings() {
        return xlSharedStrings;
    }

    /**
     * Return a Document that contains all the xl/styles.xml information
     * @return Document xl/styles.xml
     */
    public Document getXlStyles() {
        return xlStyles;
    }
    /***
     * Get the value from sharedString.xml based on the index passed by param.
     * @param sharedStrIdx Value Index
     * @return String with the value or empty String if index it's out of bounds.
     */
    public abstract String getSharedStrValue(Integer sharedStrIdx);
    /**
     * Inserts a shared string in sharedStrings.xml if the value passed is no empty.
     * @param textContext that will be placed as shared string.
     * @return int value index of the shared string if exists.
     */
    public abstract int createSharedStr(String textContext);

    /**
     * Returns the index of a specific String value if it is a sharedString, -1 if it is not a sharedString value.
     * @param value to look for in sharedStrings.xml
     * @return value index or -1 if it does not exist
     */
    public abstract int isSharedStr(String value);
    public void addSheet(XLSheet sheet){
        xlSheets.add(sheet);
    }

    public XLSheet getSheet(int index){
        return xlSheets.get(index);
    }
}
