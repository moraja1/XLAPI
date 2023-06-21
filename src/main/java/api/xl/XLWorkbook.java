package una.filesorganizeridoffice.business.api.xl;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilderFactory;
import java.nio.file.Path;
import java.util.*;

import static una.filesorganizeridoffice.business.api.xl.util.NumberUtil.isNumber;

/**
 * This class represents a xlsx workbook. This class contains main xlsx archives ass sheets, sharedString and styles and performs
 * changes into sharedString.xml file, such ass adding sharedString and returning sharedString index. It should also perform
 * other tasks such as managing styles.
 */
public final class XLWorkbook {
    private final DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
    private final String xlName;
    private final Path xlFile;
    private Document xlWorkbook;
    private Document xlSharedStrings;
    private Document xlStyles;
    private final HashMap<String, String> sheetsIdName = new HashMap();
    final List<XLSheet> xlSheets = new LinkedList<>();

    /**
     * Creates a XLWorkbook and captures its name.
     * @param xlUrl
     */
    public XLWorkbook(String xlUrl) {
        xlFile = Path.of(xlUrl);
        xlName = xlFile.getFileName().toString();
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
    public Path getXlPath() {
        return xlFile;
    }

    /**
     * Add a Sheet reference to the list that will allow Excel to map the existent sheets.
     * @param sheet_rId
     * @param sheetName
     */
    public void addSheet(String sheet_rId, String sheetName) {
        sheetsIdName.put(sheetName, sheet_rId);
    }

    /**
     * Return the sheet's name with a specific rId value.
     * @param rId ot the sheet
     * @return String with the sheet name if exists, null if not.
     */
    public String getSheetName(String rId) {
        return sheetsIdName.get(rId);
    }

    /**
     * Returns a HashMap that contains all the sheet names and rId information.
     * @return HashMap
     */
    public HashMap<String, String> getSheets() {
        return sheetsIdName;
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

    public DocumentBuilderFactory getDbf() {
        return dbf;
    }

    /***
     * Get the value from sharedString.xml based on the index passed by param.
     * @param sharedStrIdx Value Index
     * @return String with the value or empty String if index it's out of bounds.
     */
    public String getSharedStrValue(Integer sharedStrIdx) {
        String v;
        NodeList siNodes = xlSharedStrings.getElementsByTagName("si");
        if(sharedStrIdx >= 0 && sharedStrIdx < siNodes.getLength()){
            Element e = (Element) siNodes.item(sharedStrIdx);
            Element t = (Element) e.getFirstChild();
            v = t.getTextContent();
            if(t.hasAttribute("xml:space")){
                v = v.replaceAll("\\s+$", "");
            }
            return v;
        }
        return "";
    }
    private int isSharedStr(String value) {
        NodeList siNodes = xlSharedStrings.getElementsByTagName("si");
        int siNodesLength = siNodes.getLength();
        if(siNodesLength > 0){
            for(int i = 0; i < siNodesLength; i++){
                Element siTag = (Element) siNodes.item(i);
                Element tTag = (Element)siTag.getFirstChild();

                String tValue = tTag.getTextContent();
                if(tValue.equals(value)){
                    return i;
                }
            }
        }
        return -1;
    }

    /**
     * Inserts a shared string in sharedStrings.xml if the value passed is no empty.
     * @param textContext that will be placed as shared string.
     * @return int value index of the shared string if exists.
     */
    public int createSharedStr(String textContext) {
        int idx = isSharedStr(textContext);
        if(!textContext.isEmpty() && !isNumber(textContext)){
            if(idx != -1){
                return idx;
            }
            Element sstTag = (Element) xlSharedStrings.getElementsByTagName("sst").item(0);
            Element siTag = xlSharedStrings.createElement("si");
            Element tTag = xlSharedStrings.createElement("t");

            tTag.setTextContent(textContext);
            siTag.appendChild(tTag);
            sstTag.appendChild(siTag);

            String count = sstTag.getAttribute("count");
            count = String.valueOf(Integer.parseInt(count) + 1);

            sstTag.setAttribute("count", count);
            sstTag.setAttribute("uniqueCount", count);
        }
        return isSharedStr(textContext);
    }
}
