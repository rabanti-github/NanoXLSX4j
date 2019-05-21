/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Class representing a reader for the shared strings table of XLSX files
 * @author Raphael Stoeckli
 */
public class SharedStringsReader
{

    private final List<String> sharedStrings;

    /**
     * Gets the shared strings list
     * @return ArrayList of shared string entries
     */
    public List<String> getSharedStrings() {
        return sharedStrings;
    }

    /**
     * Gets whether the workbook contains shared strings
     * @return True if at least one shared string object exists in the workbook
     */
    public boolean hasElements()
    {
        return sharedStrings.size() > 0;
    }

    /**
     * Gets the value of the shared string table by its index
     * @param index Index of the stared string entry
     * @return Determined shared string value. Returns null in case of a invalid index
     */
    public String getString(int index)
    {
        if (hasElements() == false || index > sharedStrings.size() - 1)
        {
            return null;
        }
        return sharedStrings.get(index);
    }

    /**
     * Default constructor
     */
    public SharedStringsReader()
    {
        this.sharedStrings = new ArrayList<>();
    }

    /**
     * Reads the XML file form the passed stream and processes the shared strings table
     * @param stream Stream of the XML file
     * @throws IOException Throws IOException in case of an error
     */
    public void read(InputStream stream) throws IOException {
        sharedStrings.clear();
        XMLStreamReader xr;
        XMLInputFactory factory = XMLInputFactory.newFactory();
        try {
            xr = factory.createXMLStreamReader(stream);
            boolean isStringItem = false, isText = false;
            StringBuilder sb = new StringBuilder();
            String name;
            int type;
            while (xr.hasNext()) {
                type = xr.next();
                if (type == XMLStreamReader.START_ELEMENT) {
                    name = xr.getName().getLocalPart().toLowerCase();
                    if (name.equals("si") && xr.isStartElement() == true) {
                        isStringItem = true;
                    } else if (name.equals("t") && isStringItem == true) {
                        isText = true;
                    }
                }
                else if ( type == XMLStreamReader.END_ELEMENT)
                {
                    name = xr.getName().getLocalPart().toLowerCase();
                    if (name.equals("si") && isStringItem == true && isText == true) {
                        sharedStrings.add(sb.toString());
                        sb.setLength(0);
                        isText = false;
                        isStringItem = false;
                    }
                }
                else if (type == XMLStreamReader.CHARACTERS)
                {
                    if (isText)
                    {
                        sb.append(xr.getText());
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