/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.lowLevel;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.w3c.dom.Document;

/**
 * Class representing a packer to compile XLSX files
 *
 * @author Raphael Stoeckli
 */
public class Packer {

	// ### C O N S T A N T S ###

	/**
	 * Path of the main content type file (MSXML)
	 */
	private static final String CONTENT_TYPE_DOCUMENT = "[Content_Types].xml";

	// ### P R I V A T E F I E L D S ###
	private final List<String> contentTypeList;
	private final List<byte[]> dataList;
	private final List<Boolean> includeContentType;
	private final List<String> pathList;
	private final List<Relationship> relationships;
	private final XlsxWriter xlsxWriterReference;

	// ### C O N S T R U C T O R S ###

	/**
	 * Default constructor with parameter
	 *
	 * @param reference
	 *            Reference to the low level instance
	 */
	public Packer(XlsxWriter reference) {
		dataList = new ArrayList<>();
		pathList = new ArrayList<>();
		contentTypeList = new ArrayList<>();
		relationships = new ArrayList<>();
		includeContentType = new ArrayList<>();
		this.xlsxWriterReference = reference;
	}

	// ### M E T H O D S ###

	/**
	 * Adds a Part to the file
	 *
	 * @param name
	 *            Filename with relative path
	 * @param contentType
	 *            URL with information about the content type (MSXML).<br>
	 *            This information is used in the main content type file
	 * @param document
	 *            XML document to add
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.IOException
	 *             Thrown if the document could not be converted to a byte array
	 */
	public void addPart(String name, String contentType, Document document) throws ch.rabanti.nanoxlsx4j.exceptions.IOException {
		dataList.add(XlsxWriter.createBytesFromDocument(document));
		pathList.add(name);
		contentTypeList.add(contentType);
		includeContentType.add(true);
	}

	/**
	 * Adds a Part to the file
	 *
	 * @param name
	 *            Filename with relative path
	 * @param contentType
	 *            URL with information about the content type (MSXML).<br>
	 *            This information is used in the main content type file
	 * @param document
	 *            XML document to add
	 * @param includeInContentType
	 *            If true, the content type will be added in the main content type
	 *            file, otherwise not
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.IOException
	 *             Thrown if the document could not be converted to a byte array
	 */
	public void addPart(String name, String contentType, Document document, boolean includeInContentType) throws ch.rabanti.nanoxlsx4j.exceptions.IOException {
		dataList.add(XlsxWriter.createBytesFromDocument(document));
		pathList.add(name);
		contentTypeList.add(contentType);
		includeContentType.add(includeInContentType);
	}

	/**
	 * Creates the main content type file (MSXML)
	 *
	 * @return Returns the byte array to add into the compilation
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.IOException
	 *             Thrown if the document could not be converted to a byte array
	 */
	private byte[] createContentTypeDocument() throws ch.rabanti.nanoxlsx4j.exceptions.IOException {
		StringBuilder sb = new StringBuilder();
		sb.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n");
		sb.append("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />\r\n");
		sb.append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />\r\n");

		for (int i = 0; i < this.contentTypeList.size(); i++) {
			if (Boolean.FALSE.equals(this.includeContentType.get(i))) {
				continue;
			}
			sb.append("<Override PartName=\"/");
			sb.append(this.pathList.get(i));
			sb.append("\" ContentType=\"");
			sb.append(this.contentTypeList.get(i));
			sb.append("\" />\r\n");
		}
		sb.append("</Types>");
		Document d = this.xlsxWriterReference.createXMLDocument(sb.toString());
		return XlsxWriter.createBytesFromDocument(d);
	}

	/**
	 * Creates a relationship. This will be used to generate a .rels file in the
	 * compilation (MSXML)
	 *
	 * @param path
	 *            relative path and filename to rels file (e.g. _rels/.rels)
	 * @return Returns the object reference to add relationship entries
	 */
	public Relationship createRelationship(String path) {
		Relationship r = new Relationship(path);
		this.relationships.add(r);
		return r;
	}

	/**
	 * Creates a relationship file (MSXML)
	 *
	 * @param rel
	 *            Relationship object to process
	 * @return Returns the byte array to add into the compilation
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.IOException
	 *             Thrown if the document could not be converted to a byte array
	 */
	private byte[] createRelationshipDocument(Relationship rel) throws ch.rabanti.nanoxlsx4j.exceptions.IOException {
		StringBuilder sb = new StringBuilder();
		sb.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n");
		for (int i = 0; i < rel.getIdList().size(); i++) {
			sb.append("<Relationship Target=\"");
			sb.append(rel.getTargetList().get(i));
			sb.append("\" Type=\"");
			sb.append(rel.getTypeList().get(i));
			sb.append("\" Id=\"");
			sb.append(rel.getIdList().get(i));
			sb.append("\"/>\r\n");
		}
		sb.append("</Relationships>");
		Document d = this.xlsxWriterReference.createXMLDocument(sb.toString());
		return XlsxWriter.createBytesFromDocument(d);
	}

	/**
	 * Method to pack the data into a XLSX file. This is the actual compiling and
	 * writing method (to a OutputStream)
	 *
	 * @param stream
	 *            OutputStream to save the data into
	 * @throws ch.rabanti.nanoxlsx4j.exceptions.IOException
	 *             Thrown in case of a error while packing or writing
	 */
	public void pack(OutputStream stream) throws ch.rabanti.nanoxlsx4j.exceptions.IOException {
		try {
			byte[] contentTypes = createContentTypeDocument();

			ZipOutputStream out = new ZipOutputStream(new BufferedOutputStream(stream), StandardCharsets.UTF_8);
			out.setMethod(ZipOutputStream.DEFLATED);
			ZipEntry entry = new ZipEntry(CONTENT_TYPE_DOCUMENT);
			out.putNextEntry(entry);
			out.write(contentTypes, 0, contentTypes.length);
			byte[] data;
			for (int i = 0; i < this.relationships.size(); i++) {
				data = createRelationshipDocument(this.relationships.get(i));
				entry = new ZipEntry(this.relationships.get(i).getRootFolder());
				out.putNextEntry(entry);
				out.write(data, 0, data.length);
			}
			for (int i = 0; i < this.dataList.size(); i++) {
				data = this.dataList.get(i);
				entry = new ZipEntry(this.pathList.get(i));
				out.putNextEntry(entry);
				out.write(data, 0, data.length);
			}
			out.flush();
			out.close();
		}
		catch (Exception e) {
			throw new ch.rabanti.nanoxlsx4j.exceptions.IOException("There was an error while packing the file. Please see the inner exception.", e);
		}
	}

	// ### S U B C L A S S E S ###

	/**
	 * Nested class representing a relationship (MSXML)
	 */
	public static class Relationship {
		private final String rootFolder;
		private final List<String> targetList;
		private final List<String> typeList;
		private final List<String> idList;
		private int currentId;

		/**
		 * Gets the root folder of the relationship
		 *
		 * @return Root folder of the relationship
		 */
		String getRootFolder() {
			return rootFolder;
		}

		/**
		 * Gets the list of targets
		 *
		 * @return ArrayList of targets as strings
		 */
		List<String> getTargetList() {
			return targetList;
		}

		/**
		 * Gets the list of types
		 *
		 * @return ArrayList of types as strings
		 */
		List<String> getTypeList() {
			return typeList;
		}

		/**
		 * Gets the list of IDs (rId...)
		 *
		 * @return ArrayList of IDs as strings
		 */
		List<String> getIdList() {
			return idList;
		}

		/**
		 * Constructor with definition of the root folder
		 *
		 * @param path
		 *            Root folder of the relationship
		 */
		Relationship(String path) {
			this.idList = new ArrayList<>();
			this.targetList = new ArrayList<>();
			this.typeList = new ArrayList<>();
			this.rootFolder = path;
			this.currentId = 1;
		}

		/**
		 * Adds a relationship entry to the relationship
		 *
		 * @param target
		 *            Target of the entry
		 * @param type
		 *            Type of the entry
		 */
		public void addRelationshipEntry(String target, String type) {
			this.targetList.add(target);
			this.typeList.add(type);
			String id = "rId" + this.currentId;
			this.idList.add(id);
			this.currentId++;
		}
	}

}
