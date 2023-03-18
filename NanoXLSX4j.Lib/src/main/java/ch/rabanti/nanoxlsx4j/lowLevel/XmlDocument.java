/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.lowLevel;

import static javax.xml.stream.XMLStreamConstants.CHARACTERS;
import static javax.xml.stream.XMLStreamConstants.END_ELEMENT;
import static javax.xml.stream.XMLStreamConstants.START_ELEMENT;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.stream.Collectors;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;

/**
 * Class for simplified XML access without dealing with start and end elements
 * or streams. It loads an XML file from a stream and processes it in a
 * structured document, similar to the XmlDocument class in .NET (System.Xml)
 */
public class XmlDocument {

	private static final String NO_SUCH_ELEMENT_MESSAGE = "The next element cannot be returned, since out of range";

	private XmlNode documentElement;

	/**
	 * Gets the root node of the XML document
	 *
	 * @return XML node (top level)
	 */
	public XmlNode getDocumentElement() {
		return this.documentElement;
	}

	/**
	 * Loads an XML document from a stream
	 *
	 * @param stream
	 *            Input Stream
	 * @throws IOException
	 *             Throws IOException in case of an error, caught by the library
	 * @throws java.io.IOException
	 *             Throws IOException in case of a stream error
	 */
	public void load(InputStream stream) throws IOException, java.io.IOException {
		XMLStreamReader xr;
		XMLInputFactory factory = XMLInputFactory.newFactory();
		factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);

		try {
			xr = factory.createXMLStreamReader(stream);

			while (xr.hasNext()) {

				int nodeType = xr.next();
				if (nodeType == START_ELEMENT) {
					this.documentElement = XmlNode.loadXmlNode(xr);
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
	 * Class representing an iterable list of XML nodes
	 */
	public static class XmlNodeList implements Iterable<XmlNode>, Iterator<XmlNode> {

		private final ArrayList<XmlNode> items = new ArrayList<>();
		private int count = 0;

		@Override
		public boolean hasNext() {
			return this.count < items.size();
		}

		@Override
		public XmlNode next() {
			if (!hasNext()) {
				throw new NoSuchElementException(NO_SUCH_ELEMENT_MESSAGE);
			}
			return this.items.get(count++);
		}

		@Override
		public Iterator<XmlNode> iterator() {
			return new Iterator<>() {
				private int index = 0;

				@Override
				public boolean hasNext() {
					return index < items.size();
				}

				@Override
				public XmlNode next() {
					if (!hasNext()) {
						throw new NoSuchElementException(NO_SUCH_ELEMENT_MESSAGE);
					}
					return items.get(index++);
				}
			};

		}

		@Override
		public void forEach(Consumer<? super XmlNode> action) {
			Objects.requireNonNull(action);
			for (XmlNode node : this) {
				action.accept(node);
			}
		}

		/**
		 * Gets the size of XML node list
		 *
		 * @return Number of elements
		 */
		public int size() {
			return items.size();
		}

		/**
		 * Adds an XML node to the node list
		 *
		 * @param node
		 *            Item to add
		 */
		public void add(XmlNode node) {
			this.items.add(node);
		}

		/**
		 * Adds a list of XML nodes to the node list
		 *
		 * @param list
		 *            List of items
		 */
		public void addRange(List<XmlNode> list) {
			this.items.addAll(list);
		}

		/**
		 * Gets an XML node at the given index
		 *
		 * @param index
		 *            Index of the node
		 * @return XML node
		 */
		public XmlNode get(int index) {
			return this.items.get(index);
		}
	}

	/**
	 * Class representing a single XML node with possible attributes and sub-nodes
	 */
	public static class XmlNode {

		private final String name;
		private String innerText;
		private final XmlAttributeCollection attributes = new XmlAttributeCollection();
		private final XmlNodeList nodeList = new XmlNodeList();

		/**
		 * Gets the child nodes of this node
		 *
		 * @return List of XML nodes
		 */
		public XmlNodeList getChildNodes() {
			return this.nodeList;
		}

		/**
		 * Gets the (tag) name of the XML node
		 *
		 * @return Name as string
		 */
		public String getName() {
			return name;
		}

		/**
		 * Gets the inner plain text of the node, if available
		 *
		 * @return Text between the start and end tag of this XML node
		 */
		public String getInnerText() {
			return innerText;
		}

		/**
		 * Gets whether child nodes are availble in this node
		 *
		 * @return True if sub-nodes exist
		 */
		public boolean hasChildNodes() {
			return nodeList.size() > 0;
		}

		/**
		 * Constructor with definition of the node (tag) name
		 *
		 * @param name
		 *            Name of the node
		 */
		public XmlNode(String name) {
			this.name = name;
		}

		/**
		 * Gets the XML attribute of the current node by its name
		 *
		 * @param targetName
		 *            Name of the target attribute
		 * @return Attribute value as string or default value if not found (can be null)
		 */
		public String getAttribute(String targetName) {
			return getAttribute(targetName, null);
		}

		/**
		 * Gets the XML attribute of the current node by its name
		 *
		 * @param targetName
		 *            Name of the target attribute
		 * @param fallbackValue
		 *            Optional fallback value if the attribute was not found
		 * @return Attribute value as string or default value if not found (can be null)
		 */
		public String getAttribute(String targetName, String fallbackValue) {
			XmlAttribute attribute = this.attributes.getByName(targetName);
			if (attribute == null) {
				return fallbackValue;
			}
			return attribute.value;
		}

		/**
		 * Gets a list of sub-nodes with the defined name
		 *
		 * @param name
		 *            (Tag) name of the XML node
		 * @param recursively
		 *            If true, all levels are considered, otherwise only the current
		 *            level (sub-nodes of current node)
		 * @return XmlNodeList object
		 */
		public XmlNodeList getElementsByTagName(String name, boolean recursively) {
			XmlNodeList list = new XmlNodeList();
			List<XmlNode> result = this.getChildNodes().items.stream().filter(x -> x.getName().equals(name)).collect(Collectors.toList());
			list.addRange(result);
			if (recursively) {
				for (XmlNode node : this.getChildNodes()) {
					list.addRange(node.getElementsByTagName(name, true).items);
				}
			}
			return list;
		}

		/**
		 * Static method to resolve attributes and sub-nodes recursively
		 *
		 * @param reader
		 *            XML stream reader (reference)
		 * @return Resolved XML node with possible attributes and sub-nodes
		 * @throws XMLStreamException
		 *             Throws IOException in case of a stream error
		 */
		public static XmlNode loadXmlNode(XMLStreamReader reader) throws XMLStreamException {
			String elementName = reader.getName().getLocalPart();
			XmlNode node = new XmlNode(elementName);
			int attributeCount = reader.getAttributeCount();
			for (int i = 0; i < attributeCount; i++) {
				String attributeName = reader.getAttributeName(i).getLocalPart();
				String attributeValue = reader.getAttributeValue(i);
				XmlAttribute attribute = new XmlAttribute(attributeName, attributeValue);
				node.attributes.add(attribute);
			}
			StringBuilder innerText = new StringBuilder();
			while (reader.hasNext()) {

				int nodeType = reader.next();
				if (nodeType == END_ELEMENT && reader.getName().getLocalPart().equals(elementName)) {
					if (innerText.length() > 0) {
						node.innerText = innerText.toString();
					}
					return node;
				}
				else if (nodeType == START_ELEMENT) {
					XmlNode childNode = loadXmlNode(reader);
					node.nodeList.add(childNode);
				}
				else if (nodeType == CHARACTERS) {
					innerText.append(reader.getText());
				}
			}
			return node;
		}
	}

	/**
	 * Class representing an XML attribute
	 */
	public static class XmlAttribute {

		private final String name;
		private final String value;

		/**
		 * Constructor with full declaration
		 *
		 * @param name
		 *            Attribute name
		 * @param value
		 *            Attribute value
		 */
		public XmlAttribute(String name, String value) {
			this.name = name;
			this.value = value;
		}
	}

	/**
	 * Class representing an iterable list of XML attributes
	 */
	public static class XmlAttributeCollection implements Iterable<XmlAttribute>, Iterator<XmlAttribute> {

		private final ArrayList<XmlAttribute> items = new ArrayList<>();
		private int count = 0;

		@Override
		public boolean hasNext() {
			return this.count < items.size();
		}

		@Override
		public XmlAttribute next() {
			if (!hasNext()) {
				throw new NoSuchElementException(NO_SUCH_ELEMENT_MESSAGE);
			}
			return this.items.get(count++);
		}

		@Override
		public Iterator<XmlAttribute> iterator() {
			return new Iterator<>() {
				private int index = 0;

				@Override
				public boolean hasNext() {
					return index < items.size();
				}

				@Override
				public XmlAttribute next() {
					if (!hasNext()) {
						throw new NoSuchElementException(NO_SUCH_ELEMENT_MESSAGE);
					}
					return items.get(index++);
				}
			};
		}

		@Override
		public void forEach(Consumer<? super XmlAttribute> action) {
			Objects.requireNonNull(action);
			for (XmlAttribute attribute : this) {
				action.accept(attribute);
			}
		}

		/**
		 * Ads an XML attribute to the node list
		 *
		 * @param attribute
		 *            Item to add
		 */
		public void add(XmlAttribute attribute) {
			this.items.add(attribute);
		}

		/**
		 * Finds the attribute with the defined name (can be null)
		 *
		 * @param name
		 *            Name of the attribute
		 * @return XM L attribute or null if the attribute was not found
		 */
		public XmlAttribute getByName(String name) {
			return this.items.stream().filter(x -> x.name.equals(name)).findFirst().orElse(null);
		}

	}
}
