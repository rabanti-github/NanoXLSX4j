/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

/**
 * Class representing a reader to decompile XLSX files
 * @author Raphael Stoeckli
 */
public class XlsxReader
{

    private String filePath;
    private InputStream inputStream;
    private ByteArrayInputStream memoryStream;
    private final Map<Integer,WorksheetReader> worksheets;
    private WorkbookReader workbook;

    /**
     * Constructor with file path as parameter
     * @param path File path of the XLSX file to load
     */
    public XlsxReader(String path)
    {
        this.filePath = path;
        this.worksheets = new HashMap<>();
    }

    /**
     * Constructor with stream as parameter
     * @param stream Stream of the XLSX file to load
     */
    public XlsxReader(InputStream stream)
    {
        this.worksheets = new HashMap<>();
        this.inputStream = stream;
    }

    /**
     * Gets the input stream of the specified file in the archive (XLSX file)
     * @param name Name of the XML file within the XLSX file
     * @param file Zip file (XLSX)
     * @return InputStream of the specified file
     * @throws IOException  Throws IOException in case of an error
     */
    private InputStream getEntryStream(String name, ZipFile file) throws IOException
    {
        InputStream is = null;

            try
            {
                ZipEntry comparison;
                if (file != null)
                {
                    comparison = file.getEntry(name);
                    is = file.getInputStream(comparison);
                }
                else
                {
                    memoryStream.reset();
                    ZipInputStream zs = new ZipInputStream(memoryStream);
                    while ((comparison = zs.getNextEntry()) != null)
                    {
                        if (comparison.getName().equals(name))
                        {
                            is = zs;
                            break;
                        }
                    }
                }
                return is;
            }
            catch (Exception ex)
            {
                throw new IOException("LoadException", "There was an error while extracting a stream from a XLSX file. Please see the inner exception:", ex);
            }
    }

    /**
     * Reads the XLSX file from a file path or a file stream
     * @throws IOException  Throws IOException in case of an error
     */
    public void read() throws IOException
    {
        ZipFile zf = null;
        try
        {
            ZipEntry entry;
            if (inputStream == null || Helper.isNullOrEmpty(filePath) == false)
            {
                zf = new ZipFile(this.filePath);
            }
            else if (inputStream != null)
            {
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                byte[] buffer = new byte[1024];
                for(int i = inputStream.read(buffer); i != -1; i = inputStream.read(buffer))
                {
                    os.write(buffer, 0, i);
                }
                inputStream.close();
                memoryStream = new ByteArrayInputStream(os.toByteArray());
            }
            else
            {
                throw new IOException("LoadException", "No valid stream or file path was provided to open");
            }
            InputStream stream;
            SharedStringsReader sharedStrings = new SharedStringsReader();
            stream = getEntryStream("xl/sharedStrings.xml", zf);
            sharedStrings.read(stream);

            this.workbook = new WorkbookReader();
            stream = getEntryStream("xl/workbook.xml", zf);
            this.workbook.read(stream);

            int worksheetIndex = 1;
            String name, nameTemplate;
            WorksheetReader wr;
            nameTemplate = "sheet" + Integer.toString(worksheetIndex) + ".xml";
            name = "xl/worksheets/" + nameTemplate;
            for (int i = 0; i < this.workbook.getWorksheetDefinitions().size(); i++)
            {
                stream = getEntryStream(name, zf);
                wr = new WorksheetReader(sharedStrings, nameTemplate, worksheetIndex);
                wr.read(stream);
                this.worksheets.put(worksheetIndex - 1, wr);
                worksheetIndex++;
                nameTemplate = "sheet" + Integer.toString(worksheetIndex) + ".xml";
                name = "xl/worksheets/" + nameTemplate;
            }
        }
        catch (Exception ex)
        {
            throw new IOException("LoadException", "There was an error while reading an XLSX file. Please see the inner exception:", ex);
        }
        finally
        {
            if (memoryStream != null)
            {
                try {
                    memoryStream.close();
                } catch (java.io.IOException ex) {
                    throw new IOException("XMLStreamException", "The XML memory stream could not be closed. Please see the inner exception:", ex);
                }
            }
        }
    }

    /**
     * Resolves the workbook with all worksheets from the loaded file
     * @return Workbook object
     */
    public Workbook getWorkbook()
    {
        Workbook wb = new Workbook(false);
        WorksheetReader wr;
        Worksheet ws;
        for(int i = 0; i < this.worksheets.size(); i++)
        {
            wr = this.worksheets.get(i);
            ws = new Worksheet(this.workbook.getWorksheetDefinitions().get(i + 1), i+1, wb);
            for (Map.Entry<String, Cell> cell : wr.getData().entrySet())
            {
                ws.addCell(cell.getValue(), cell.getKey());
            }
            wb.addWorksheet(ws);
        }
        return wb;
    }

}