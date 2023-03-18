/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.ImportOptions;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;

/**
 * Class representing a reader for the shared strings table of XLSX files
 *
 * @author Raphael Stoeckli
 */
public class SharedStringsReader {

	private final List<String> sharedStrings;
	private boolean capturePhoneticCharacters = false;
	private List<PhoneticInfo> phoneticsInfo = null;

	/**
	 * Gets whether the workbook contains shared strings
	 *
	 * @return True if at least one shared string object exists in the workbook
	 */
	public boolean hasElements() {
		return !sharedStrings.isEmpty();
	}

	/**
	 * Gets the value of the shared string table by its index
	 *
	 * @param index
	 *            Index of the stared string entry
	 * @return Determined shared string value. Returns null in case of a invalid
	 *         index
	 */
	public String getString(int index) {
		if (!hasElements() || index > sharedStrings.size() - 1 || index < 0) {
			return null;
		}
		return sharedStrings.get(index);
	}

	/**
	 * Constructor with parameters
	 *
	 * @param importOptions
	 *            Import options instance
	 */
	public SharedStringsReader(ImportOptions importOptions) {
		if (importOptions != null) {
			this.capturePhoneticCharacters = importOptions.isEnforcePhoneticCharacterImport();
			if (this.capturePhoneticCharacters) {
				this.phoneticsInfo = new ArrayList<>();
			}
		}
		this.sharedStrings = new ArrayList<>();
	}

	/**
	 * Reads the XML file form the passed stream and processes the shared strings
	 * table
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
			StringBuilder sb = new StringBuilder();
			for (XmlDocument.XmlNode node : xr.getDocumentElement().getChildNodes()) {
				if (node.getName().equalsIgnoreCase("si")) {
					sb = new StringBuilder();
					getTextToken(node, sb);
					if (this.capturePhoneticCharacters) {
						this.sharedStrings.add(processPhoneticCharacters(sb));
					}
					else {
						this.sharedStrings.add(sb.toString());
					}
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
	 * Function collects text tokens recursively in case of a split by formatting
	 *
	 * @param node
	 *            Root node to process
	 * @param sb
	 *            StringBuilder reference
	 */
	private void getTextToken(XmlDocument.XmlNode node, StringBuilder sb) {
		if (node.getName().equalsIgnoreCase("rPh")) {
			if (this.capturePhoneticCharacters) {
				if (node.hasChildNodes() &&
						node.getChildNodes().get(0).getName().equalsIgnoreCase("t") &&
						!Helper.isNullOrEmpty(node.getChildNodes().get(0).getInnerText())) {
					String start = node.getAttribute("sb");
					String end = node.getAttribute("eb");
					this.phoneticsInfo.add(new PhoneticInfo(node.getChildNodes().get(0).getInnerText(), start, end));
				}
			}
			return;
		}

		if (node.getName().equalsIgnoreCase("t") && !Helper.isNullOrEmpty(node.getInnerText())) {
			// Reproduces the new line behavior of Excel
			String text = node.getInnerText().replace("\r\n", "\n").replace("\n", "\r\n");
			sb.append(text);
		}
		if (node.hasChildNodes()) {
			for (XmlDocument.XmlNode childNode : node.getChildNodes()) {
				getTextToken(childNode, sb);
			}
		}
	}

	/**
	 * Function to add determined phonetic tokens
	 *
	 * @param sb
	 *            Original StringBuilder
	 * @return Text with added phonetic characters (after particular characters, in
	 *         brackets)
	 */
	private String processPhoneticCharacters(StringBuilder sb) {
		if (this.phoneticsInfo.isEmpty()) {
			return sb.toString();
		}
		String text = sb.toString();
		StringBuilder sb2 = new StringBuilder();
		int currentTextIndex = 0;
		for (PhoneticInfo info : this.phoneticsInfo) {
			sb2.append(text, currentTextIndex, info.getStartIndex() + info.length);
			sb2.append("(").append(info.getValue()).append(")");
			currentTextIndex = info.getStartIndex() + info.getLength();
		}
		sb2.append(text.substring(currentTextIndex));

		phoneticsInfo.clear();
		return sb2.toString();
	}

	/**
	 * Class to represent a phonetic transcription of character sequence.
	 *
	 * @implNote Invalid values will lead to a crash. The specifications requires a
	 *           start index, an end index and a value
	 */
	private static class PhoneticInfo {
		private final String value;
		private final int startIndex;
		private final int length;

		/**
		 * Gets the transcription value
		 *
		 * @return Transcription (phonetic characters)
		 */
		public String getValue() {
			return value;
		}

		/**
		 * Gets the absolute start index within the original string
		 *
		 * @return Zero-based start index
		 */
		public int getStartIndex() {
			return startIndex;
		}

		/**
		 * Gets the number of characters of the original string that are described by
		 * this transcription token
		 *
		 * @return Number of characters
		 */
		public int getLength() {
			return length;
		}

		/**
		 * Constructor with parameters
		 *
		 * @param value
		 *            Transcription value
		 * @param start
		 *            Absolute start index as string
		 * @param end
		 *            Absolute end index as string
		 */
		public PhoneticInfo(String value, String start, String end) {
			this.value = value;
			this.startIndex = Integer.parseInt(start);
			this.length = Integer.parseInt(end) - this.startIndex;
		}
	}

}