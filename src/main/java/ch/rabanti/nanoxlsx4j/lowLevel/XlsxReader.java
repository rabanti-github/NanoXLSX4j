/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.ImportOptions;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Style;

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
 *
 * @author Raphael Stoeckli
 */
public class XlsxReader {

    private String filePath;
    private InputStream inputStream;
    private ByteArrayInputStream memoryStream;
    private final Map<Integer, WorksheetReader> worksheets;
    private WorkbookReader workbook;
    private ImportOptions importOptions;
    private StyleReaderContainer styleReaderContainer;

    /**
     * Constructor with stream and import options as parameter
     *
     * @param stream  Stream of the XLSX file to load
     * @param options Import options to override the automatic approach of the reader. See {@link ImportOptions} for information about import options
     */
    public XlsxReader(InputStream stream, ImportOptions options) {
        this.worksheets = new HashMap<>();
        this.inputStream = stream;
        this.importOptions = options;
    }

    /**
     * Constructor with file path and import options as parameter
     *
     * @param path    File path of the XLSX file to load
     * @param options Import options to override the automatic approach of the reader. See {@link ImportOptions} for information about import options
     */
    public XlsxReader(String path, ImportOptions options) {
        this.filePath = path;
        this.worksheets = new HashMap<>();
        this.importOptions = options;
    }

    /**
     * Gets the input stream of the specified file in the archive (XLSX file)
     *
     * @param name Name of the XML file within the XLSX file
     * @param file Zip file (XLSX)
     * @return InputStream of the specified file
     * @throws IOException Throws IOException in case of an error
     */
    private InputStream getEntryStream(String name, ZipFile file) throws IOException {
        InputStream is = null;

        try {
            ZipEntry comparison;
            if (file != null) {
                comparison = file.getEntry(name);
                is = file.getInputStream(comparison);
            } else {
                memoryStream.reset();
                ZipInputStream zs = new ZipInputStream(memoryStream);
                while ((comparison = zs.getNextEntry()) != null) {
                    if (comparison.getName().equals(name)) {
                        is = zs;
                        break;
                    }
                }
            }
            if (is == null) {
                throw new IOException("The entry '" + name + "' is missing in the file");
            }
            return is;
        } catch (Exception ex) {
            throw new IOException("There was an error while extracting a stream from a XLSX file. Please see the inner exception:", ex);
        }
    }

    /**
     * Reads the XLSX file from a file path or a file stream
     *
     * @throws IOException Throws IOException in case of an error
     */
    public void read() throws IOException, java.io.IOException {
        ZipFile zf = null;
        try {
            if (inputStream == null && !Helper.isNullOrEmpty(filePath)) {
                zf = new ZipFile(this.filePath);
            } else if (inputStream != null) {
                ByteArrayOutputStream os = new ByteArrayOutputStream();
                byte[] buffer = new byte[1024];
                for (int i = inputStream.read(buffer); i != -1; i = inputStream.read(buffer)) {
                    os.write(buffer, 0, i);
                }
                inputStream.close();
                memoryStream = new ByteArrayInputStream(os.toByteArray());
            } else {
                throw new IOException("No valid stream or file path was provided to open");
            }
            InputStream stream;
            SharedStringsReader sharedStrings = new SharedStringsReader(importOptions);
            stream = getEntryStream("xl/sharedStrings.xml", zf);
            sharedStrings.read(stream);

            StyleReader styleReader = new StyleReader();
            stream = getEntryStream("xl/styles.xml", zf);
            styleReader.read(stream);
            this.styleReaderContainer = styleReader.getStyleReaderContainer();

            this.workbook = new WorkbookReader();
            stream = getEntryStream("xl/workbook.xml", zf);
            this.workbook.read(stream);

            int worksheetIndex = 1;
            WorksheetReader wr;
            String nameTemplate = "sheet" + worksheetIndex + ".xml";
            String name = "xl/worksheets/" + nameTemplate;
            for (Map.Entry<Integer, WorkbookReader.WorksheetDefinition> definition : workbook.getWorksheetDefinitions().entrySet()) {
                stream = getEntryStream(name, zf);
                wr = new WorksheetReader(sharedStrings, styleReaderContainer, importOptions);
                wr.read(stream);
                this.worksheets.put(definition.getKey(), wr);
                worksheetIndex++;
                nameTemplate = "sheet" + worksheetIndex + ".xml";
                name = "xl/worksheets/" + nameTemplate;
            }
        } catch (Exception ex) {
            throw new IOException("There was an error while reading an XLSX file. Please see the inner exception:", ex);
        } finally {
            if (memoryStream != null) {
                memoryStream.close();
            }
        }
    }

    /**
     * Resolves the workbook with all worksheets from the loaded file
     *
     * @return Workbook object
     */
    public Workbook getWorkbook() {
        Workbook wb = new Workbook(false);
        wb.setImportState(true);
        Worksheet ws;
        for (Map.Entry<Integer, WorksheetReader> reader : this.worksheets.entrySet()) {
            WorkbookReader.WorksheetDefinition definition = workbook.getWorksheetDefinitions().get(reader.getKey());
            ws = new Worksheet(definition.getWorksheetName(), definition.getSheetId(), wb);
            ws.setHidden(definition.isHidden());
            if (reader.getValue().getAutoFilterRange() != null) {
                ws.setAutoFilter(reader.getValue().getAutoFilterRange().StartAddress.Column, reader.getValue().getAutoFilterRange().EndAddress.Column);
            }
            if (reader.getValue().getDefaultColumnWidth() != null){
                ws.setDefaultColumnWidth(reader.getValue().getDefaultColumnWidth());
            }
            if (reader.getValue().getDefaultRowHeight() != null){
                ws.setDefaultRowHeight(reader.getValue().getDefaultRowHeight());
            }
            if (reader.getValue().getSelectedCells() != null){
                ws.setSelectedCells(reader.getValue().getSelectedCells());
            }
            for(Range range : reader.getValue().getMergedCells()){
                ws.mergeCells(range);
            }
            for(Map.Entry<Worksheet.SheetProtectionValue, Integer> sheetProtection : reader.getValue().getWorksheetProtection().entrySet())
            {
                ws.getSheetProtectionValues().add(sheetProtection.getKey());
            }
            if (!reader.getValue().getWorksheetProtection().isEmpty())
            {
                ws.setUseSheetProtection(true);
            }
            for(Map.Entry<Integer,WorksheetReader.RowDefinition> row : reader.getValue().getRows().entrySet())
            {
                if (row.getValue().isHidden())
                {
                    ws.addHiddenRow(row.getKey());
                }
                if (row.getValue().getHeight() != null)
                {
                    ws.setRowHeight(row.getKey(), row.getValue().getHeight());
                }
            }
            for(Column column : reader.getValue().getColumns()){
                if (column.getWidth() != Worksheet.DEFAULT_COLUMN_WIDTH){
                    ws.setColumnWidth(column.getColumnAddress(), column.getWidth());
                }
                if (column.isHidden()){
                    ws.addHiddenColumn(column.getNumber());
                }
            }
            for (Map.Entry<String, Cell> cell : reader.getValue().getData().entrySet()) {
                if (reader.getValue().getStyleAssignment().containsKey(cell.getKey())) {
                    Style style = styleReaderContainer.getStyle(reader.getValue().getStyleAssignment().get(cell.getKey()), true);
                    if (style != null) {
                        cell.getValue().setStyle(style);
                    }
                }
                ws.addCell(cell.getValue(), cell.getKey());
            }
            wb.addWorksheet(ws);
        }
        wb.setImportState(false);
        return wb;
    }

}
