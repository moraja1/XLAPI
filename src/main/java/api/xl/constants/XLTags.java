package api.xl.constants;

import api.xl.base.XLFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.util.function.Supplier;

public class XLTags {
    private static final DocumentBuilderFactory dbFactory;
    private static final DocumentBuilder dBuilder;
    private static final Document doc;
    static{
        dbFactory = DocumentBuilderFactory.newInstance();
        try {
            dBuilder = dbFactory.newDocumentBuilder();
        } catch (ParserConfigurationException e) {
            throw new RuntimeException(e);
        }
        doc = dBuilder.newDocument();
        doc.setXmlStandalone(true);
    }
    public static final class ShredStr{
        public static final String SST = "sst";
        public static final String SST_C = "count";
        public static final String SST_UC = "uniqueCount";
        public static final String NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public static final String SI = "si";
        public static final String T = "t";
        public static final String T_XML = "xml";
        public static final String XML_SPACE = "space";
        public static final String T_XML_VAL = "preserve";
        public static Document defaultSharedStr() {
            Element SST = defaultSST();
            doc.appendChild(SST);
            return (Document) doc.cloneNode(true);
        }
        public static Element defaultSST(){
            Element sstTag = doc.createElementNS(NS, SST);
            sstTag.setAttribute(SST_C, "0");
            sstTag.setAttribute(SST_UC, "0");
            return sstTag;
        }
    }

    public static final class Sheet{
        public static final String ROW = "row";
        public static final String R_ATTR = "r";
        public static final String SPANS_ATTR = "spans";
        public static final String ROW_NS = "x14ac";
        public static final String ROW_QN = "dyDescent";
        public static final String CELL = "c";
        public static final String VAL = "v";
    }
}
