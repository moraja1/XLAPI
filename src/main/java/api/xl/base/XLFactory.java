package api.xl.base;

import api.xl.XLSheet;
import api.xl.XLWorkbook;
import api.xl.exceptions.XLFactoryException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.Enumeration;
import java.util.Objects;
import java.util.function.Predicate;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import static api.xl.util.NumberUtil.isNumber;

public final class XLFactory {

    private static Predicate<Document> exists = Objects::isNull;

    /**
     * Return a XLWorkbook object if the xlsx file passed exists, it's openable and readable, and it's a xlsx file.
     * @param xlsx file
     * @return XLWorkbook if file passed is accepted, null if not.
     * @throws XLFactoryException when XLWorkbook can not be created.
     */
    public static XLWorkbook buildWorkbook(File xlsx) throws XLFactoryException {
        XLWorkbook w;
        if(!xlsx.exists() || xlsx.isDirectory() || !xlsx.isFile() || !xlsx.getName().endsWith(".xlsx")){
            return null;
        }
        try {
            w = new Workbook(xlsx);
        } catch (IOException e) {
            throw new XLFactoryException("Unable to create XLWorkbook, please check the filepath and try again.");
        }
        //Obtaining xml basic files
        ZipFile zipFile = w.getXlFile();
        Enumeration<? extends ZipEntry> entries = zipFile.entries();
        while(entries.hasMoreElements()){
            ZipEntry entry = entries.nextElement();
            //System.out.println(entry.getName());
        /*if (entry.getName().equals(SHARED_STRINGS)) {
            xlWorkbook.setXlSharedStrings(createDocument(zipFile, entry));
        }else if (entry.getName().equals(STYLES)) {
            xlWorkbook.setXlStyles(createDocument(zipFile, entry));
        }else if (entry.getName().equals(WORKBOOK)) {
            xlWorkbook.setXlWorkbook(createDocument(zipFile, entry));
        }else if(entry.getName().contains("sheet")){
            //Set sheets in xlWorkbook
            Document workbook = xlWorkbook.getXlWorkbook();
            NodeList sheets = workbook.getElementsByTagName("sheet");
        }*/
        }

        /*for (int i = 0; i < sheets.getLength(); i++) {
            Element sheet = (Element) sheets.item(i);
            String sheetName = sheet.getAttributeNode("name").getValue();
            String sheet_rId = sheet.getAttributeNode("r:id").getValue();
            //xlWorkbook.addSheet(sheet_rId, sheetName);
        }*/
        if((exists.test(w.getXlWorkbook())) || exists.test(w.getXlTheme()) ||
               exists.test(w.getXlStyles())  || exists.test(w.getXlSharedStrings())){
            return null;
        }
        return w;
    }
    private static Document createDocument(ZipFile z,ZipEntry entry) throws IOException, ParserConfigurationException, SAXException {
        InputStream inputStream = z.getInputStream(entry);
        return DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(inputStream);
    }

    /**
     * Creates a Document of the sheet with the filename "sheet" plus the index passed by param plus ".xml" located in
     * the Workbook file passed by Param.
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

    private static final class Workbook extends XLWorkbook {
        /**
         * Creates a XLWorkbook and captures its name.
         * @param xlsx
         */
        public Workbook(File xlsx) throws IOException {
            super(xlsx);
        }
        @Override
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

        @Override
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

        @Override
        public int isSharedStr(String value) {
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
    }
}