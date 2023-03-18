/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipInputStream;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Column;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.ImportOptions;
import ch.rabanti.nanoxlsx4j.Range;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.Worksheet;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;

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
	private MetaDataReader metaDataReader;
	private final ImportOptions importOptions;
	private StyleReaderContainer styleReaderContainer;

	/**
	 * Constructor with stream and import options as parameter
	 *
	 * @param stream
	 *            Stream of the XLSX file to load
	 * @param options
	 *            Import options to override the automatic approach of the reader.
	 *            See {@link ImportOptions} for information about import options
	 */
	public XlsxReader(InputStream stream, ImportOptions options) {
		this.worksheets = new HashMap<>();
		this.inputStream = stream;
		this.importOptions = options;
	}

	/**
	 * Constructor with file path and import options as parameter
	 *
	 * @param path
	 *            File path of the XLSX file to load
	 * @param options
	 *            Import options to override the automatic approach of the reader.
	 *            See {@link ImportOptions} for information about import options
	 */
	public XlsxReader(String path, ImportOptions options) {
		this.filePath = path;
		this.worksheets = new HashMap<>();
		this.importOptions = options;
	}

	/**
	 * Gets the input stream of the specified file in the archive (XLSX file)
	 *
	 * @param name
	 *            Name of the XML file within the XLSX file
	 * @param file
	 *            Zip file (XLSX)
	 * @return InputStream of the specified file
	 * @throws IOException
	 *             Throws IOException in case of an error
	 */
	private InputStream getEntryStream(String name, ZipFile file) throws IOException {
		return getEntryStream(name, file, true);
	}

	/**
	 * Gets the input stream of the specified file in the archive (XLSX file)
	 *
	 * @param name
	 *            Name of the XML file within the XLSX file
	 * @param file
	 *            Zip file (XLSX)
	 * @param throwExceptionIfAbsent If true, an exception will be thrown if the stream could not be found, otherwise null is returned
	 * @return InputStream of the specified file or null if the stream was not found and parameter throwExceptionIfAbsent was set to false
	 * @throws IOException
	 *             Throws IOException in case of an error and if throwExceptionIfAbsent is true
	 */
	private InputStream getEntryStream(String name, ZipFile file , boolean throwExceptionIfAbsent) throws IOException {
		InputStream is = null;

		try {
			ZipEntry comparison;
			if (file != null) {
				comparison = file.getEntry(name);
				is = file.getInputStream(comparison);
			}
			else {
				memoryStream.reset();
				ZipInputStream zs = new ZipInputStream(memoryStream);
				while ((comparison = zs.getNextEntry()) != null) {
					if (comparison.getName().equals(name)) {
						is = zs;
						break;
					}
				}
			}
			if (is == null && throwExceptionIfAbsent) {
				throw new IOException("The entry '" + name + "' is missing in the file");
			}
			return is;
		}
		catch (Exception ex) {
			throw new IOException("There was an error while extracting a stream from a XLSX file. Please see the inner exception:", ex);
		}
	}

	/**
	 * Reads the XLSX file from a file path or a file stream
	 *
	 * @throws IOException
	 *             Throws IOException in case of an error
	 */
	public void read() throws IOException, java.io.IOException {
		ZipFile zf = null;
		try {
			if (inputStream == null && !Helper.isNullOrEmpty(filePath)) {
				zf = new ZipFile(this.filePath);
			}
			else if (inputStream != null) {
				ByteArrayOutputStream os = new ByteArrayOutputStream();
				byte[] buffer = new byte[1024];
				for (int i = inputStream.read(buffer); i != -1; i = inputStream.read(buffer)) {
					os.write(buffer, 0, i);
				}
				inputStream.close();
				memoryStream = new ByteArrayInputStream(os.toByteArray());
			}
			else {
				throw new IOException("No valid stream or file path was provided to open");
			}
			InputStream stream;
			SharedStringsReader sharedStrings = new SharedStringsReader(importOptions);
			stream = getEntryStream("xl/sharedStrings.xml", zf, false);
			if (stream != null) {
				sharedStrings.read(stream);
			}
			StyleRepository.getInstance().setImportInProgress(true);
			StyleReader styleReader = new StyleReader();
			stream = getEntryStream("xl/styles.xml", zf);
			styleReader.read(stream);
			this.styleReaderContainer = styleReader.getStyleReaderContainer();
			StyleRepository.getInstance().setImportInProgress(false);

			this.workbook = new WorkbookReader();
			stream = getEntryStream("xl/workbook.xml", zf);
			this.workbook.read(stream);

			metaDataReader = new MetaDataReader();
			stream = getEntryStream("docProps/app.xml", zf);
			this.metaDataReader.ReadAppData(stream);
			stream = getEntryStream("docProps/core.xml", zf);
			this.metaDataReader.ReadCoreData(stream);

			int worksheetIndex = 1;
			WorksheetReader wr;
			String nameTemplate = "sheet" + worksheetIndex + ".xml";
			String name = "xl/worksheets/" + nameTemplate; // default
			RelationshipReader relationships = new RelationshipReader();
			stream = getEntryStream("xl/_rels/workbook.xml.rels", zf);
			relationships.read(stream);
			for (Map.Entry<Integer, WorkbookReader.WorksheetDefinition> definition : workbook.getWorksheetDefinitions().entrySet()) {
				Optional<RelationshipReader.RelationShip> relationship = relationships.getRelationships().stream().filter(r -> r.getId().equals(definition.getValue().getRelId())).findFirst();
				if (relationship.isPresent()){
					// relationship resolution
					name = relationship.get().getTarget();
				}
				stream = getEntryStream(name, zf);
				wr = new WorksheetReader(sharedStrings, styleReaderContainer, importOptions);
				wr.read(stream);
				this.worksheets.put(definition.getKey(), wr);
				// fallback resolution
				worksheetIndex++;
				nameTemplate = "sheet" + worksheetIndex + ".xml";
				name = "xl/worksheets/" + nameTemplate;
			}
			if (this.worksheets.isEmpty()){
				throw new IOException("No worksheet was found in the workbook");
			}
		}
		catch (Exception ex) {
			throw new IOException("There was an error while reading an XLSX file. Please see the inner exception:", ex);
		}
		finally {
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
			if (reader.getValue().getDefaultColumnWidth() != null) {
				ws.setDefaultColumnWidth(reader.getValue().getDefaultColumnWidth());
			}
			if (reader.getValue().getDefaultRowHeight() != null) {
				ws.setDefaultRowHeight(reader.getValue().getDefaultRowHeight());
			}
			if (reader.getValue().getSelectedCells() != null) {
				for(Range range : reader.getValue().getSelectedCells()){
					ws.addSelectedCells(range);
				}
			}
			for (Range range : reader.getValue().getMergedCells()) {
				ws.mergeCells(range);
			}
			for (Map.Entry<Worksheet.SheetProtectionValue, Integer> sheetProtection : reader.getValue().getWorksheetProtection().entrySet()) {
				ws.getSheetProtectionValues().add(sheetProtection.getKey());
			}
			if (!reader.getValue().getWorksheetProtection().isEmpty()) {
				ws.setUseSheetProtection(true);
			}
			if (!Helper.isNullOrEmpty(reader.getValue().getWorksheetProtectionHash())) {
				ws.setSheetProtectionPasswordHash(reader.getValue().getWorksheetProtectionHash());
			}
			for (Map.Entry<Integer, WorksheetReader.RowDefinition> row : reader.getValue().getRows().entrySet()) {
				if (row.getValue().isHidden()) {
					ws.addHiddenRow(row.getKey());
				}
				if (row.getValue().getHeight() != null) {
					ws.setRowHeight(row.getKey(), row.getValue().getHeight());
				}
			}
			for (Column column : reader.getValue().getColumns()) {
				if (column.getWidth() != Worksheet.DEFAULT_COLUMN_WIDTH) {
					ws.setColumnWidth(column.getColumnAddress(), column.getWidth());
				}
				if (column.isHidden()) {
					ws.addHiddenColumn(column.getNumber());
				}
			}
			for (Map.Entry<String, Cell> cell : reader.getValue().getData().entrySet()) {
				if (reader.getValue().getStyleAssignment().containsKey(cell.getKey())) {
					Style style = styleReaderContainer.getStyle(reader.getValue().getStyleAssignment().get(cell.getKey()));
					if (style != null) {
						cell.getValue().setStyle(style);
					}
				}
				ws.addCell(cell.getValue(), cell.getKey());
			}
			if (reader.getValue().getPaneSplitValue() != null) {
				WorksheetReader.PaneDefinition pane = reader.getValue().getPaneSplitValue();
				if (pane.getFrozenState()) {
					if (pane.isYSplitDefined() && !pane.isXSplitDefined()) {
						ws.setHorizontalSplit(pane.getPaneSplitRowIndex(), pane.getFrozenState(), pane.getTopLeftCell(), pane.getActivePane());
					}
					if (!pane.isYSplitDefined() && pane.isXSplitDefined()) {
						ws.setVerticalSplit(pane.getPaneSplitColumnIndex(), pane.getFrozenState(), pane.getTopLeftCell(), pane.getActivePane());
					}
					else if (pane.isYSplitDefined() && pane.isXSplitDefined()) {
						ws.setSplit(pane
								.getPaneSplitColumnIndex(), pane.getPaneSplitRowIndex(), pane.getFrozenState(), pane.getTopLeftCell(), pane.getActivePane());
					}
				}
				else {
					if (pane.isYSplitDefined() && !pane.isXSplitDefined()) {
						ws.setHorizontalSplit(pane.getPaneSplitHeight(), pane.getTopLeftCell(), pane.getActivePane());
					}
					if (!pane.isYSplitDefined() && pane.isXSplitDefined()) {
						ws.setVerticalSplit(pane.getPaneSplitWidth(), pane.getTopLeftCell(), pane.getActivePane());
					}
					else if (pane.isYSplitDefined() && pane.isXSplitDefined()) {
						ws.setSplit(pane.getPaneSplitWidth(), pane.getPaneSplitHeight(), pane.getTopLeftCell(), pane.getActivePane());
					}
				}
			}
			wb.addWorksheet(ws);
		}
		if (!styleReaderContainer.getMruColors().isEmpty()) {
			for (String color : styleReaderContainer.getMruColors()) {
				wb.addMruColor(color);
			}
		}
		wb.setHidden(workbook.isHidden());
		wb.setSelectedWorksheet(workbook.getSelectedWorksheet());
		if (workbook.isProtected()) {
			wb.setWorkbookProtection(workbook.isProtected(), workbook.isLockWindows(), workbook.isLockStructure(), null);
			wb.setWorkbookProtectionPasswordHash(workbook.getPasswordHash());
		}
		wb.getWorkbookMetadata().setApplication(metaDataReader.getApplication());
		wb.getWorkbookMetadata().setApplicationVersion(metaDataReader.getApplicationVersion());
		wb.getWorkbookMetadata().setCreator(metaDataReader.getCreator());
		wb.getWorkbookMetadata().setCategory(metaDataReader.getCategory());
		wb.getWorkbookMetadata().setCompany(metaDataReader.getCompany());
		wb.getWorkbookMetadata().setContentStatus(metaDataReader.getContentStatus());
		wb.getWorkbookMetadata().setDescription(metaDataReader.getDescription());
		wb.getWorkbookMetadata().setHyperlinkBase(metaDataReader.getHyperlinkBase());
		wb.getWorkbookMetadata().setKeywords(metaDataReader.getKeywords());
		wb.getWorkbookMetadata().setManager(metaDataReader.getManager());
		wb.getWorkbookMetadata().setSubject(metaDataReader.getSubject());
		wb.getWorkbookMetadata().setTitle(metaDataReader.getTitle());
		wb.setImportState(false);
		return wb;
	}

}
