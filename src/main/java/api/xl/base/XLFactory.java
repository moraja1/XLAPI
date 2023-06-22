package api.xl.base;

import api.xl.constants.XLTags;
import api.xl.exceptions.XLFactoryException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.w3c.dom.ls.DOMImplementationLS;
import org.w3c.dom.ls.LSSerializer;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.Enumeration;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import static api.xl.constants.XLPaths.*;
import static api.xl.util.XLNumberUtil.isNumber;

/**
 * This class contains the necessary methods to create, open and save workbooks as well as some important implementations
 */
public final class XLFactory {

    private static final Predicate<Document> exists = Objects::nonNull;

    /**
     * Return a XLWorkbook object if the xlsx file passed exists, it's openable and readable, and it's a xlsx file.
     * @param xlsx file
     * @return XLWorkbook if file passed is accepted, null if not.
     * @throws XLFactoryException when XLWorkbook can not be created.
     */
    public static XLWorkbook openWorkbook(File xlsx) throws XLFactoryException {
        if (!xlsx.exists() || xlsx.isDirectory() || !xlsx.isFile() || !xlsx.getName().endsWith(".xlsx")) {
            return null;
        }
        final XLWorkbook w;
        final ZipFile zf;
        try {
            zf = new ZipFile(xlsx);
        } catch (IOException e) {
            throw new XLFactoryException("Unable to open Workbook, please check the filepath and try again.");
        }
        //Load required Documents
        final Document workbookDoc;
        final Document sharedStrDoc;
        final Document stylesDoc;

        ZipEntry workbookEntry = zf.getEntry(WORKBOOK);
        workbookDoc = createDocument(zf, workbookEntry);

        ZipEntry sharedStringsEntry = zf.getEntry(SHARED_STRINGS);
        if(sharedStringsEntry == null){
            sharedStrDoc = XLTags.ShredStr.defaultSharedStr();
        }else{
            sharedStrDoc = createDocument(zf, sharedStringsEntry);
        }

        ZipEntry stylesEntry = zf.getEntry(STYLES);
        stylesDoc = createDocument(zf, stylesEntry);

        w = new XLWorkbookImp(zf, workbookDoc, sharedStrDoc, stylesDoc);

        Enumeration<? extends ZipEntry> entries = zf.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = entries.nextElement();
            if (entry.getName().contains("sheet")) {
                //Set sheets in xlWorkbook
                Document workbook = w.getXlWorkbook();
                NodeList sheets = workbook.getElementsByTagName("sheet");
                for (int i = 0; i < sheets.getLength(); i++) {
                    Element sheet = (Element) sheets.item(i);
                    String sheetName = sheet.getAttribute("name");
                    String sheet_rId = sheet.getAttribute("r:id");
                    String sheetId = sheet.getAttribute("sheetId");
                    if(entry.getName().equals(SHEET.apply(sheetId))){
                        XLSheet xlSheet = new XLSheet(w, createDocument(zf, entry), sheetName, sheet_rId, sheetId);
                        w.addSheet(xlSheet);
                    }
                }
            }
        }

        //Verify results
        if ((exists.test(w.getXlWorkbook())) ||
                exists.test(w.getXlStyles()) || exists.test(w.getXlSharedStrings())) {
            return w;
        }
        return null;
    }

    /**
     * Creates and returns a new empty workbook
     * @return XLWorkbook in an empty state
     * @throws XLFactoryException if there is an error opening the new workbook.
     */
    public static XLWorkbook newWorkbook() throws XLFactoryException {
        URL url = XLWorkbook.class.getResource("../Book1.xlsx");
        XLWorkbook w;
        try {
            File xlsx = new File(url.toURI());
            w = XLFactory.openWorkbook(xlsx);
        } catch (URISyntaxException | XLFactoryException e) {
            throw new XLFactoryException("Unable to create a new Workbook, please use openWorkbook method and pass a valid xlsx file.");
        }
        return  w;
    }

    private static Document createDocument(ZipFile z, ZipEntry entry) {
        InputStream inputStream = null;
        try {
            inputStream = z.getInputStream(entry);
            return DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(inputStream);
        } catch (IOException | ParserConfigurationException | SAXException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Creates a Document of the sheet with the filename "sheet" plus the index passed by param plus ".xml" located in
     * the Workbook file passed by Param.
     *
     * @param w XLWorkbook
     * @param i Integer
     * @return XLSheet if the sheet with the name exists, or null if not.
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    public static XLSheet buildSheet(XLWorkbook w, Integer i) throws IOException, ParserConfigurationException, SAXException {
        /*ZipFile zipFile = new ZipFile(w.getXlPath().toFile());
        ZipEntry entry;
        String entryName = SHEET.apply(i.toString());
        entry = zipFile.getEntry(entryName);
        if(entry != null){

            XLSheet xlSheet = new XLSheet(w, createDocument(zipFile, entry), entryName);
            w.xlSheets.add(xlSheet);
            return xlSheet;
        }*/
        return null;
    }

    /**
     * This method saves the changes made on any book passed by param. This also receives the path where the new xlsx file
     * will be located.
     *
     * @param xlWorkbook
     * @param filepath
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws TransformerException
     * @throws IOException
     */
    public static void saveWorkbook(XLWorkbook xlWorkbook, String filepath) throws ParserConfigurationException, SAXException, TransformerException, IOException {
        //Create a xlsx in the temp directory where the information will be placed.
        /*Path newXL = Paths.get(filepath);
        if(!Files.exists(newXL)){
            Files.copy(xlWorkbook.getXlPath(), newXL);
        }

        //Iterate over entries and save them into the new xlsx. Verifies the xml files and changes the last xml for the new one
        try(ZipFile zipFile = new ZipFile(xlWorkbook.getXlPath().toFile());
            ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(newXL.toFile()))){
            boolean entryProcessed = false;
            byte[] buffer = new byte[1024];
            int bytesRead;
            byte[] documentBytes = null;
            Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while(entries.hasMoreElements()){
                ZipEntry entry = entries.nextElement();
                String entryName = entry.getName();
                ZipEntry newEntry;

                if (entryName.equals(SHARED_STRINGS)) {
                    documentBytes = convertDocumentToBytes(xlWorkbook.getXlSharedStrings());
                    entryProcessed = true;
                }else if (entryName.equals(STYLES)) {
                    documentBytes = convertDocumentToBytes(xlWorkbook.getXlStyles());
                    entryProcessed = true;
                }else if (entryName.equals(WORKBOOK)) {
                    documentBytes = convertDocumentToBytes(xlWorkbook.getXlWorkbook());
                    entryProcessed = true;
                }else if(entryName.contains("sheet")){
                    for (int i = 1; i <= xlWorkbook.xlSheets.size(); i++) {
                        if(entry.getName().equals(SHEET.apply(String.valueOf(i)))){
                            documentBytes = convertDocumentToBytes(xlWorkbook.xlSheets.get(i-1).getXlSheet());
                            entryProcessed = true;
                        }
                    }
                }
                newEntry = new ZipEntry(entryName);
                zipOut.putNextEntry(newEntry);

                if(entryProcessed) {
                    if(documentBytes != null){
                        zipOut.write(documentBytes);
                    }
                    entryProcessed = false;
                }else{
                    InputStream inputStream = zipFile.getInputStream(entry);
                    while ((bytesRead = inputStream.read(buffer)) != -1) {
                        zipOut.write(buffer, 0, bytesRead);
                    }
                    inputStream.close();
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }*/

    }

    private static void transformXML(Document document, File entryFile) throws TransformerException, IOException {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        DOMSource source = new DOMSource(document);
        FileWriter writer = new FileWriter(entryFile);
        StreamResult result = new StreamResult(writer);
        transformer.transform(source, result);
    }

    private static byte[] convertDocumentToBytes(Document document) {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.transform(new DOMSource(document), new StreamResult(outputStream));

            return outputStream.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /***
     * Print in console any Node. Test purposes.
     * @param node xml Node to be printed.
     */
    public static void toStringNode(Element node){
        Document document = node.getOwnerDocument();
        DOMImplementationLS domImplLS = (DOMImplementationLS) document.getImplementation();
        LSSerializer serializer = domImplLS.createLSSerializer();
        serializer.getDomConfig().setParameter("format-pretty-print", true);
        String str = serializer.writeToString(node);
        System.out.println(str);
    }

    private static class XLWorkbookImp extends XLWorkbook {

        /**
         * Creates a XLWorkbook and captures its name.
         *
         * @param xlsx
         * @param xlWorkbook
         * @param xlSharedStrings
         * @param xlStyles
         */
        public XLWorkbookImp(ZipFile xlsx, Document xlWorkbook, Document xlSharedStrings, Document xlStyles) {
            super(xlsx, xlWorkbook, xlSharedStrings, xlStyles);
        }

        /***
         * Get the value from sharedString.xml based on the index passed by param.
         * @param sharedStrIdx Value Index
         * @return String with the value or empty String if index it's out of bounds.
         */
        @Override
        public String getSharedStrValue(Integer sharedStrIdx){
            String v;
            NodeList siNodes = xlSharedStrings.getElementsByTagName("si");
            if (sharedStrIdx >= 0 && sharedStrIdx < siNodes.getLength()) {
                Element e = (Element) siNodes.item(sharedStrIdx);
                Element t = (Element) e.getFirstChild();
                v = t.getTextContent();
                if (t.hasAttribute("xml:space")) {
                    v = v.replaceAll("\\s+$", "");
                }
                return v;
            }
            return "";
        }
        /**
         * Inserts a shared string in sharedStrings.xml if the value passed is no empty.
         * @param textContext that will be placed as shared string.
         * @return int value index of the shared string if exists.
         */
        @Override
        public int createSharedStr(String textContext){
            int idx = isSharedStr(textContext);
            if (!textContext.isEmpty() && !isNumber(textContext)) {
                if (idx != -1) {
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

        /**
         * Returns the index of a specific String value if it is a sharedString, -1 if it is not a sharedString value.
         * @param value to look for in sharedStrings.xml
         * @return value index or -1 if it does not exist
         */
        @Override
        public int isSharedStr(String value) {
            NodeList siNodes = xlSharedStrings.getElementsByTagName("si");
            int siNodesLength = siNodes.getLength();
            if (siNodesLength > 0) {
                for (int i = 0; i < siNodesLength; i++) {
                    Element siTag = (Element) siNodes.item(i);
                    Element tTag = (Element) siTag.getFirstChild();

                    String tValue = tTag.getTextContent();
                    if (tValue.equals(value)) {
                        return i;
                    }
                }
            }
            return -1;
        }
    }
}