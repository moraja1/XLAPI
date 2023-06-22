package api.xl.base;

import org.w3c.dom.Document;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.zip.ZipFile;

/**
 * This class represents a xlsx workbook. This class contains main xlsx archives ass sheets, sharedString and styles and performs
 * changes into sharedString.xml file, such ass adding sharedString and returning sharedString index. It should also perform
 * other tasks such as managing styles.
 */
public abstract class XLWorkbook {
    private final String xlName;
    private final ZipFile xlFile;
    protected Document xlWorkbook;
    protected Document xlSharedStrings;
    protected Document xlStyles;
    final List<XLSheet> xlSheets = new LinkedList<>();
    /**
     * Creates a XLWorkbook and captures its name.
     * @param xlsx
     */
    public XLWorkbook(File xlsx) throws IOException {
        xlFile = new ZipFile(xlsx);
        xlName = xlsx.getName();
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
     * @return
     */
    public Document getXlWorkbook() {
        return xlWorkbook;
    }
    public void setXlWorkbook(Document xlWorkbook) {
        this.xlWorkbook = xlWorkbook;
    }
    public Document getXlSharedStrings() {
        return xlSharedStrings;
    }
    public void setXlSharedStrings(Document xlSharedStrings) {
        this.xlSharedStrings = xlSharedStrings;
    }
    public Document getXlStyles() {
        return xlStyles;
    }
    public void setXlStyles(Document xlStyles) {
        this.xlStyles = xlStyles;
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
}
