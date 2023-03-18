/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

/**
 * Class representing a reader to decompile a workbook in an XLSX files
 *
 * @author Raphael Stoeckli
 */
public class WorkbookReader {
	private final Map<Integer, WorksheetDefinition> worksheetDefinitions;
	private boolean hidden;
	private int selectedWorksheet;
	private boolean wbProtected;
	private boolean lockWindows;
	private boolean lockStructure;
	private String passwordHash;

	/**
	 * Hashmap of worksheet definitions. The key is the worksheet number and the
	 * value is a WorksheetDefinition object with name, hidden state and other
	 * information
	 *
	 * @return Hashmap of number-name tuples
	 */
	public Map<Integer, WorksheetDefinition> getWorksheetDefinitions() {
		return worksheetDefinitions;
	}

	/**
	 * Hidden state of the workbook
	 *
	 * @return True if hidden
	 */
	public boolean isHidden() {
		return hidden;
	}

	/**
	 * Selected worksheet of the workbook
	 *
	 * @return Index of the worksheet
	 */
	public int getSelectedWorksheet() {
		return selectedWorksheet;
	}

	/**
	 * Protection state of the workbook
	 *
	 * @return True if protected
	 */
	public boolean isProtected() {
		return wbProtected;
	}

	/**
	 * Lock state of the windows
	 *
	 * @return True if locked
	 */
	public boolean isLockWindows() {
		return lockWindows;
	}

	/**
	 * Lock state of the structural elements
	 *
	 * @return True if locked
	 */
	public boolean isLockStructure() {
		return lockStructure;
	}

	/**
	 * Password hash if available
	 *
	 * @return Hash as string
	 */
	public String getPasswordHash() {
		return passwordHash;
	}

	/**
	 * Default constructor
	 */
	public WorkbookReader() {
		this.worksheetDefinitions = new HashMap<>();
	}

	/**
	 * Reads the XML file form the passed stream and processes the workbook
	 * information
	 *
	 * @param stream
	 *            Stream of the XML file
	 * @throws IOException
	 *             Throws IOException in case of an error
	 */
	public void read(InputStream stream) throws IOException, java.io.IOException {
		try {
			XmlDocument xr = new XmlDocument();
			xr.load(stream);
			for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
				if (node.getName().equalsIgnoreCase("sheets") && node.hasChildNodes()) {
					getWorksheetInformation(node.getChildNodes());
				}
				else if (node.getName().equalsIgnoreCase("bookViews") && node.hasChildNodes()) {
					getViewInformation(node.getChildNodes());
				}
				else if (node.getName().equalsIgnoreCase("workbookProtection")) {
					getProtectionInformation(node);
				}

			}
		}
		catch (Exception ex) {
			throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
		}
		finally {
			if (stream != null) {
				stream.close();
			}
		}
	}

	/**
	 * Gets the workbook protection information
	 *
	 * @param node
	 *            Root node to check
	 */
	private void getProtectionInformation(XmlDocument.XmlNode node) {
		this.wbProtected = true;
		String attribute = node.getAttribute("lockWindows");
		if (attribute != null) {
			int value = ReaderUtils.parseBinaryBoolean(attribute);
			this.lockWindows = value == 1;
		}
		attribute = node.getAttribute("lockStructure");
		if (attribute != null) {
			int value = ReaderUtils.parseBinaryBoolean(attribute);
			this.lockStructure = value == 1;
		}
		attribute = node.getAttribute("workbookPassword");
		if (attribute != null) {
			this.passwordHash = attribute;
		}
	}

	/**
	 * Gets the workbook view information
	 *
	 * @param nodes
	 *            View nodes to check
	 */
	private void getViewInformation(XmlDocument.XmlNodeList nodes) {
		for (XmlDocument.XmlNode node : nodes) {
			String attribute = node.getAttribute("visibility");
			if (attribute != null && attribute.equalsIgnoreCase("hidden")) {
				this.hidden = true;
			}
			attribute = node.getAttribute("activeTab");
			if (attribute != null && !attribute.isEmpty()) {
				this.selectedWorksheet = Integer.parseInt(attribute);
			}
		}
	}

	/**
	 * Gets the worksheet information
	 *
	 * @param nodes
	 *            Sheet nodes to check
	 * @throws IOException
	 *             Thrown if the workbook information could not be determined
	 */
	private void getWorksheetInformation(XmlDocument.XmlNodeList nodes) throws IOException {
		for (XmlDocument.XmlNode node : nodes) {
			if (node.getName().equalsIgnoreCase("sheet")) {
				try {
					String sheetName = node.getAttribute("name", "worksheet1");
					int id = Integer.parseInt(node.getAttribute("sheetId")); // Default will rightly throw an exception
					String relId = node.getAttribute("id"); // Note: namespace 'r' is stripped (r:id)
					String state = node.getAttribute("state");
					boolean hidden = state != null && state.equalsIgnoreCase("hidden");
					WorksheetDefinition definition = new WorksheetDefinition(id, sheetName, relId);
					definition.setHidden(hidden);
					worksheetDefinitions.put(id, definition);
				}
				catch (Exception e) {
					throw new IOException("The workbook information could not be resolved. Please see the inner exception:", e);
				}
			}
		}
	}

	// ### S U B - C L A S S E S ###

	/**
	 * Class for worksheet Mata-data on import
	 */
	public static class WorksheetDefinition {
		private final String worksheetName;
		private final int sheetId;
		private String relId;
		private boolean hidden;

		/**
		 * Sets the hidden state of the worksheet
		 *
		 * @param hidden
		 *            True if hidden
		 */
		public void setHidden(boolean hidden) {
			this.hidden = hidden;
		}

		/**
		 * Gets the hidden state of the worksheet
		 *
		 * @return True if hidden
		 */
		public boolean isHidden() {
			return hidden;
		}

		/**
		 * Gets the worksheet name
		 *
		 * @return Name as string
		 */
		public String getWorksheetName() {
			return worksheetName;
		}

		/**
		 * Internal worksheet ID
		 *
		 * @return Intenal id
		 */
		public int getSheetId() {
			return sheetId;
		}

		/**
		 * Gets the reference ID
		 * @return Reference ID as string
		 */
		public String getRelId() {
			return relId;
		}

		/**
		 * Default constructor with parameters
		 *
		 * @param worksheetName
		 *            Worksheet name
		 * @param sheetId
		 *            Internal ID
		 * @param relId Relationship ID
		 */
		public WorksheetDefinition(int sheetId, String worksheetName, String relId) {
			this.sheetId = sheetId;
			this.worksheetName = worksheetName;
			this.relId = relId;
		}
	}

}
