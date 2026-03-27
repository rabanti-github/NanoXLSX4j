/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal.writers;

// NOTE: ch.rabanti.nanoxlsx4j.Workbook and ch.rabanti.nanoxlsx4j.Metadata are currently
// in the nanoxlsx4j-lib (Lib) module.  The Writer module will be able to compile this class
// once either:
//   (a) Workbook / Metadata are migrated to nanoxlsx4j-core, OR
//   (b) The Writer pom.xml declares a compile-time dependency on nanoxlsx4j-lib.
// This is expected as part of the ongoing v3 source migration.
import ch.rabanti.nanoxlsx4j.Metadata;
import ch.rabanti.nanoxlsx4j.Workbook;

import ch.rabanti.nanoxlsx4j.interfaces.writer.IBaseWriter;
import ch.rabanti.nanoxlsx4j.interfaces.writer.IPluginWriter;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxPlugin;
import ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable;
import ch.rabanti.nanoxlsx4j.registry.PluginUUID;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.xml.XmlElement;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;

/**
 * Writer for the core metadata XML part ({@code docProps/core.xml}).
 * <p>
 * Java equivalent of the .NET internal class
 * {@code NanoXLSX.Internal.Writers.MetadataCoreWriter}.
 * </p>
 *
 * <h3>Plugin wiring</h3>
 * <p>
 * This class is a <em>Replacing Plugin</em> registered under
 * {@link PluginUUID#META_DATA_CORE_WRITER}.  It is discovered at runtime via
 * the {@link java.util.ServiceLoader} mechanism: the fully-qualified class name
 * must be listed in
 * {@code META-INF/services/ch.rabanti.nanoxlsx4j.registry.NanoXlsxPluggable}
 * inside the Writer JAR.
 * </p>
 *
 * <h3>XmlElement migration note</h3>
 * <p>
 * {@link #getXmlElement()} currently returns a stub because
 * {@link XmlElement} is not yet fully implemented.  Until the
 * {@code XmlElement} layer is complete, callers inside the writer orchestrator
 * should use the package-private {@link #getXmlString()} accessor to obtain
 * the generated XML as a plain string.
 * </p>
 *
 * @implNote Framework-internal class. Not intended for use outside the writer module.
 * @author Raphael Stoeckli
 * @see PluginUUID#META_DATA_CORE_WRITER
 */
@NanoXlsxPlugin(plugInUuid = PluginUUID.META_DATA_CORE_WRITER)
public class MetadataCoreWriter implements NanoXlsxPluggable, IPluginWriter {

    // ### C O N S T A N T S ###

    private static final String NAMESPACE_CP =
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static final String NAMESPACE_DC       = "http://purl.org/dc/elements/1.1/";
    private static final String NAMESPACE_DCTERMS  = "http://purl.org/dc/terms/";
    private static final String NAMESPACE_DCMITYPE = "http://purl.org/dc/dcmitype/";
    private static final String NAMESPACE_XSI      = "http://www.w3.org/2001/XMLSchema-instance";

    // ### P R I V A T E  F I E L D S ###

    private Workbook workbook;
    private XmlElement xmlElement;

    /**
     * The built XML string — package-private for use by the writer orchestrator
     * until {@link XmlElement} is fully implemented.
     */
    private String xmlString;

    // ### G E T T E R S  &  S E T T E R S ###

    /**
     * Gets the workbook instance associated with this writer
     *
     * @return the workbook (typed as {@code Workbook} during the v3 migration phase)
     */
    @Override
    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * Sets the workbook instance for this writer
     *
     * @param workbook the workbook to use
     */
    @Override
    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * Gets the main XML element produced by {@link #execute()}.
     *
     * @return stub {@link XmlElement} — full implementation pending XmlElement migration
     */
    @Override
    public XmlElement getXmlElement() {
        return xmlElement;
    }

    // ### C O N S T R U C T O R S ###

    /**
     * Default constructor — required for plugin instantiation via {@link java.util.ServiceLoader}
     */
    public MetadataCoreWriter() {
    }

    // ### M E T H O D S ###

    /**
     * Initialises this writer from the supplied base-writer context (interface implementation)
     *
     * @param baseWriter base writer providing workbook, styles and shared-string access
     */
    @Override
    public void init(IBaseWriter baseWriter) {
        this.workbook = baseWriter.getWorkbook();
    }

    /**
     * Builds the core-properties XML content and stores it internally
     * (interface implementation).
     * <p>
     * The generated XML corresponds to the {@code docProps/core.xml} part of
     * an XLSX package.  Null or empty metadata fields are silently omitted.
     * </p>
     */
    @Override
    public void execute() {
        // Cast is safe when called from the NanoXLSX4j write pipeline.
        // Will compile once Workbook is available in the Writer module's classpath.
        Workbook wb = (Workbook) workbook;
        Metadata md = wb.getWorkbookMetadata();

        StringBuilder content = new StringBuilder();
        if (md != null) {
            appendXmlTag(content, md.getTitle(),         "title",          "dc");
            appendXmlTag(content, md.getSubject(),       "subject",        "dc");
            appendXmlTag(content, md.getCreator(),       "creator",        "dc");
            appendXmlTag(content, md.getCreator(),       "lastModifiedBy", "cp");
            appendXmlTag(content, md.getKeywords(),      "keywords",       "cp");
            appendXmlTag(content, md.getDescription(),   "description",    "dc");
        }

        Calendar cal = new GregorianCalendar();
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
        df.setCalendar(cal);
        Date now = cal.getTime();
        String time = df.format(now);

        content.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:created>");
        content.append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">").append(time).append("</dcterms:modified>");

        if (md != null) {
            appendXmlTag(content, md.getCategory(),      "category",      "cp");
            appendXmlTag(content, md.getContentStatus(), "contentStatus", "cp");
        }

        this.xmlString =
                "<cp:coreProperties" +
                " xmlns:cp=\""       + NAMESPACE_CP       + "\"" +
                " xmlns:dc=\""       + NAMESPACE_DC       + "\"" +
                " xmlns:dcterms=\""  + NAMESPACE_DCTERMS  + "\"" +
                " xmlns:dcmitype=\"" + NAMESPACE_DCMITYPE + "\"" +
                " xmlns:xsi=\""      + NAMESPACE_XSI      + "\">" +
                content +
                "</cp:coreProperties>";

        // TODO: build a proper XmlElement tree once XmlElement supports child elements
        // and namespace management.  For now the stub satisfies the interface contract.
        this.xmlElement = new XmlElement();
    }

    /**
     * Returns the XML generated by {@link #execute()} as a plain string.
     * <p>
     * Package-private accessor for the writer orchestrator to consume until
     * the full {@link XmlElement} implementation is in place.
     * </p>
     *
     * @return core-properties XML string, or {@code null} if {@link #execute()} has not been called
     */
    String getXmlString() {
        return xmlString;
    }

    // ### P R I V A T E  H E L P E R S ###
    // Duplicated from XlsxWriter for now; will be extracted to a shared XML
    // utility in a later migration step.

    /**
     * Appends a namespaced XML element with a text value to the given builder.
     * Does nothing if {@code value} is null or empty.
     *
     * @param sb        builder to append to
     * @param value     text content; skipped when null or empty
     * @param tagName   local tag name
     * @param nameSpace namespace prefix, or null/empty for no prefix
     */
    private void appendXmlTag(StringBuilder sb, String value, String tagName, String nameSpace) {
        if (ParserUtils.isNullOrEmpty(value)) {
            return;
        }
        boolean hasNoNs = ParserUtils.isNullOrEmpty(nameSpace);
        sb.append('<');
        if (!hasNoNs) {
            sb.append(nameSpace).append(':');
        }
        sb.append(tagName).append('>');
        sb.append(escapeXmlChars(value));
        sb.append("</");
        if (!hasNoNs) {
            sb.append(nameSpace).append(':');
        }
        sb.append(tagName).append('>');
    }

    /**
     * Escapes characters that are illegal or must be encoded in XML text content.
     * Illegal control characters are replaced with a space; {@code <}, {@code >},
     * and {@code &} are replaced with their entity references.
     *
     * @param input input string to process
     * @return escaped string; returns {@code input} unchanged if no escaping is required
     */
    private static String escapeXmlChars(String input) {
        int len = input.length();
        List<Integer> illegalCharacters = new ArrayList<>(len);
        List<Integer> characterTypes    = new ArrayList<>(len);
        char c;
        for (int i = 0; i < len; i++) {
            c = input.charAt(i);
            if ((c < 0x9) || (c > 0xA && c < 0xD) || (c > 0xD && c < 0x20)
                    || (c > 0xD7FF && c < 0xE000) || (c > 0xFFFD)) {
                illegalCharacters.add(i);
                characterTypes.add(0);   // control character → space
            }
            else if (c == 0x3C) {        // <
                illegalCharacters.add(i);
                characterTypes.add(1);
            }
            else if (c == 0x3E) {        // >
                illegalCharacters.add(i);
                characterTypes.add(2);
            }
            else if (c == 0x26) {        // &
                illegalCharacters.add(i);
                characterTypes.add(3);
            }
        }
        if (illegalCharacters.isEmpty()) {
            return input;
        }
        StringBuilder sb = new StringBuilder(len);
        int lastIndex = 0;
        int count = illegalCharacters.size();
        for (int i = 0; i < count; i++) {
            int j    = illegalCharacters.get(i);
            int type = characterTypes.get(i);
            sb.append(input, lastIndex, j);
            switch (type) {
                case 0: sb.append(' ');      break;
                case 1: sb.append("&lt;");   break;
                case 2: sb.append("&gt;");   break;
                case 3: sb.append("&amp;");  break;
                default: break;
            }
            lastIndex = j + 1;
        }
        sb.append(input.substring(lastIndex));
        return sb.toString();
    }

}
