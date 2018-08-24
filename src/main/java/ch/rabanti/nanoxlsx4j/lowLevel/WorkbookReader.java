/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exception.IOException;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * Class representing a reader to decompile a workbook in an XLSX files
 * @author Raphael Stoeckli
 */
public class WorkbookReader
{
    private final Map<Integer, String> worksheetDefinitions;

    /**
     * Hashmap of worksheet definitions. The key is the worksheet number and the value is the worksheet name
     * @return Hashmap of number-name tuples
     */
    public Map<Integer, String> getWorksheetDefinitions() {
        return worksheetDefinitions;
    }

    /**
     * Default constructor
     */
    public WorkbookReader()
    {
        this.worksheetDefinitions = new HashMap<>();
    }

    /**
     * Reads the XML file form the passed stream and processes the workbook information
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException
    {
        XMLStreamReader xr;
        XMLInputFactory factory = XMLInputFactory.newFactory();
        int nodeType, id;
        String name, sheetName, sheetId;
        try {
            xr = factory.createXMLStreamReader(stream);
            while (xr.hasNext()) {
                nodeType = xr.next();
                if (nodeType == XMLStreamReader.START_ELEMENT)
                {
                    name = xr.getName().getLocalPart().toLowerCase();
                    if (name.equals("sheet"))
                    {
                        sheetName = xr.getAttributeValue(null, "name");
                        sheetId = xr.getAttributeValue(null, "sheetId");
                        id = Integer.parseInt(sheetId);
                        this.worksheetDefinitions.put(id, sheetName);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new IOException("XMLStreamException", "The XML entry could not be read from the input stream. Please see the inner exception:", ex);
        }
        finally
        {
            if (stream != null)
            {
                try {
                    stream.close();
                } catch (java.io.IOException ex) {
                    throw new IOException("XMLStreamException", "The XML entry stream could not be closed. Please see the inner exception:", ex);
                }
            }
        }
    }


}
