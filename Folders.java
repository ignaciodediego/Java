import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/*
10. Folders
Implement a function folderNames, which accepts a string containing an XML file that specifies folder structure
and returns all folder names that start with startingLetter. The XML format is given in the example below.
For example, for the letter 'u' and an XML file:
    <?xml version="1.0" encoding="UTF-8"?>
    <folder name="c">
        <folder name="program files">
            <folder name="uninstall information" />
        </folder>
        <folder name="users" />
    </folder>
the function should return a collection with items "uninstall information" and "users" (in any order).
 */
public class Folders {
    public static Collection<String> folderNames(String xml, char startingLetter) throws Exception {
        List<String> list = new ArrayList<>();
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder db = dbf.newDocumentBuilder();
        Document document = db.parse(new InputSource(new StringReader(xml)));
        NodeList nodeList = document.getElementsByTagName("folder");
        for (int x = 0, size = nodeList.getLength(); x < size; x++) {
            String name = nodeList.item(x).getAttributes().getNamedItem("name").getNodeValue();
            if (name.startsWith(String.valueOf(startingLetter))) {
                list.add(name);
            }
        }

        return list;
    }

    public static void main(String[] args) throws Exception {
        String xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                        "<folder name=\"c\">" +
                        "<folder name=\"program files\">" +
                        "<folder name=\"uninstall information\" />" +
                        "</folder>" +
                        "<folder name=\"users\" />" +
                        "</folder>";

        Collection<String> names = folderNames(xml, 'u');
        for (String name : names)
            System.out.println(name);
    }
}