package api.xl.base;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import javax.xml.parsers.ParserConfigurationException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import static api.xl.util.XLConverter.cellTypetoString;
import static api.xl.util.XLDateUtil.isDate;
import static api.xl.util.XLDateUtil.toDate;
import static api.xl.util.XLNumberUtil.isNumber;
import static api.xl.util.XLNumberUtil.isScientificNotation;

public class XLSheet {
    protected final XLWorkbook xlWorkbook;
    protected final Document xlSheet;
    private final String name;
    private final String rId;
    private final String sheetId;
    protected final List<String> ignoreColumnCases = new ArrayList<>();

    /***
     * Creates a XLSheet
     * @param xlWorkbook
     * @param document
     */
    public XLSheet(XLWorkbook xlWorkbook, Document document, String name, String rId, String sheetId) {
        this.xlWorkbook = xlWorkbook;
        this.xlSheet = document;
        this.name = name;
        this.rId = rId;
        this.sheetId = sheetId;
    }

    /***
     * Returns the XLRow of a specific row number with all its XLCell values if the row exists, else, a base XLRow is returned.
     * This method return null if the index of the row passed if less than zero.
     * @param idx int
     * @return XLRow
     */
    public XLRow getRow(int idx){
        if(idx > 0) {
            //Cell memory space
            XLRow row = new XLRow(idx);
            XLCell<?> xlCell;
            String cellValue;
            String cellColumn;
            Integer cellRow;
            //Get Row Element
            Element rowTag = getRowElement(idx);
            //Obtains the row number
            if(rowTag != null){
                String rAttr = rowTag.getAttribute("r");
                cellRow = Integer.valueOf(rAttr);
                //Compares row number
                if (cellRow == idx) {
                    //Obtains the cells
                    NodeList cells = rowTag.getElementsByTagName("c");
                    for (int j = 0; j < cells.getLength(); j++) {
                        Element cell = (Element) cells.item(j);
                        //Prepare cell information
                        String cellPoint = cell.getAttribute("r");
                        cellColumn = cellPoint.replace(String.valueOf(cellRow), "");
                        //Ignore Column Cases
                        if (ignoreColumnCases.contains(cellColumn)) {
                            continue;
                        }
                        //Obtains index of sharedString.xml or null if it's not a sharedString
                        Integer sharedStrIdx = getSharedIdx(cell);
                        if (sharedStrIdx != null) {
                            //Obtain sharedString.xml value based on the index
                            cellValue = xlWorkbook.getSharedStrValue(sharedStrIdx);
                        } else {
                            Element v = (Element) cell.getFirstChild();
                            //Obtains the cell value
                            if (v != null) {
                                cellValue = v.getTextContent();
                            } else {
                                cellValue = "";
                            }
                        }
                        //Creates de cell with proper value Type
                        Integer intValue;
                        LocalDate dateValue;
                        String stringValue;
                        if (isScientificNotation(cellValue)) {
                            intValue = new BigDecimal(cellValue).intValue();
                            xlCell = new XLCell<Integer>(cellColumn, cellRow, intValue);
                        } else if (isDate(cellValue)) {
                            dateValue = toDate(cellValue);
                            xlCell = new XLCell<LocalDate>(cellColumn, cellRow, dateValue);
                        } else if (cellValue.matches("\\d*") && !cellValue.isEmpty()) {
                            intValue = Integer.valueOf(cellValue);
                            xlCell = new XLCell<Integer>(cellColumn, cellRow, intValue);
                        } else {
                            stringValue = cellValue;
                            xlCell = new XLCell<String>(cellColumn, cellRow, stringValue);
                        }
                        row.addXlCell(xlCell);
                        //Ends if its adult
                        if (xlCell.getValue().equals("Mayor de edad")) j = cells.getLength();
                    }
                }
            }
            return row;
        }
        return null;
    }

    /**
     * This method receives a XLRow and paste it into the sheet .xml file.
     * Note: This method will change the value of the cell that exists in the row number you are pasting if there is any.
     * If there is not a row already in that row number, then it will create the new row.
     * @param row with the cells that will be placed in the sheet.
     */
    public void pasteRow(XLRow row){
        int rowIdx = row.getRowNum();
        if(rowIdx > 0){
            Element rowTag = getRowElement(rowIdx);
            //si el row no existe en el xml creo el row
            if(rowTag == null){
                rowTag = xlSheet.createElement("row");
                rowTag.setAttribute("r", String.valueOf(rowIdx));
                rowTag.setAttribute("spans", "1:2");
                rowTag.setAttributeNS("x14ac", "dyDescent", "0.25");
            }
            //Recorro cada celda del XLRow
            for(int i = 0; i < row.getCellCount(); i++) {
                XLCell<?> xlCell = row.getCell(i);
                Element cellTag;
                String rValue = xlCell.getColumnName().concat(String.valueOf(rowIdx));
                cellTag = getCell(rowTag, rValue);

                //Obtenemos el valor que vamos a almacenar en la celda
                String textContext = cellTypetoString(xlCell.getValue());

                //Si no es numero lo creo en sharedString.xml y obtengo su indice
                int sharedStrIdx = -1;
                if(!isNumber(textContext)){
                    sharedStrIdx = xlWorkbook.createSharedStr(textContext);
                    textContext = String.valueOf(sharedStrIdx);
                }
                if(sharedStrIdx != -1){
                    cellTag.setAttribute("t", "s");
                }
                if(!textContext.equals("-1")){
                    //Busco o creo el tag v y almaceno el textContext
                    Element vTag = (Element) cellTag.getFirstChild();
                    if(vTag == null){
                        vTag = xlSheet.createElement("v");
                    }
                    vTag.setTextContent(textContext);
                    cellTag.appendChild(vTag);
                }
            }
        }
    }

    protected Element getCell(Element rowTag, String rValue) {
        NodeList xmlCells = rowTag.getElementsByTagName("c");
        Element xmlCell;

        for(int i = 0; i < xmlCells.getLength(); i++) {
            xmlCell = (Element) xmlCells.item(i);

            String rAttr = xmlCell.getAttribute("r");
            if(rAttr.equals(rValue)){
                return xmlCell;
            }
        }
        //Si la celda que voy a copiar no existe en el xml, creo una celda sin formato y se la inserto al row del xml.
        xmlCell = xlSheet.createElement("c");
        xmlCell.setAttribute("r", rValue);
        rowTag.appendChild(xmlCell);
        return xmlCell;
    }

    protected Element getRowElement(int idx) {
        NodeList rows = xlSheet.getElementsByTagName("row");
        int rowsCant = rows.getLength();
        if(idx <= rowsCant) {
            return (Element) rows.item(idx - 1);
        }
        return null;
    }

    /***
     * Obtains index of sharedString.xml or null if it's not a sharedString
     * @param e Element
     * @return Integer if it is a sharedString or null if no
     */
    protected Integer getSharedIdx(Element e) {
        if(e.hasAttribute("t")){
           String tValue = e.getFirstChild().getFirstChild().getNodeValue();
           return Integer.valueOf(tValue);
        }
        return null;
    }

    public void addIgnoreColumnCase(String column) {
        ignoreColumnCases.add(column);
    }

    public void clearIgnoreColumnCases() {
        if (!ignoreColumnCases.isEmpty()){
            ignoreColumnCases.clear();
        }
    }

    /**
     * Returns the name of the xml file that contains the sheet's information
     * @return String
     */
    public String getName() {
        return name;
    }

    public Document getXlSheet() {
        return xlSheet;
    }
}