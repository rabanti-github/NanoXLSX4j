//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.registry;

/**
 * String constants for every plugin slot defined in NanoXLSX4j.
 * <p>
 * These UUIDs are used to register and look up plugins via {@link PluginLoader}.
 * The values are identical to those in the .NET reference implementation
 * ({@code NanoXLSX.Core.Registry.PluginUUID}) to maintain cross-platform
 * consistency.
 * </p>
 *
 * <h3>Categories</h3>
 * <ul>
 *   <li><b>Writers</b> – replacing-plugin slots for each XLSX write component.</li>
 *   <li><b>Readers</b> – replacing-plugin slots for each XLSX read component.</li>
 *   <li><b>Writer queues</b> – queue-plugin slots executed around the write pass.</li>
 *   <li><b>Writer inline queues</b> – queue-plugin slots executed immediately after
 *       a specific writer component.</li>
 *   <li><b>Reader queues</b> – queue-plugin slots executed around the read pass.</li>
 *   <li><b>Reader inline queues</b> – queue-plugin slots executed immediately after
 *       a specific reader component.</li>
 *   <li><b>Entities</b> – replacing-plugin slots for shared workbook entities.</li>
 * </ul>
 *
 * @see NanoXlsxPlugin
 * @see NanoXlsxQueuePlugin
 * @see PluginLoader
 */
public final class PluginUUID {

    private PluginUUID() {
        // Utility class — not instantiable.
    }

    // -------------------------------------------------------------------------
    // WRITERS
    // -------------------------------------------------------------------------

    /** Replacing plugin: writes the password / encryption information. */
    public static final String PASSWORD_WRITER          = "8106E566-60D6-45DB-BF87-33AB3882C019";

    /** Replacing plugin: writes the workbook XML part ({@code workbook.xml}). */
    public static final String WORKBOOK_WRITER          = "D4272E3A-AC56-4524-9B9F-7B1448DF536B";

    /** Replacing plugin: writes individual worksheet XML parts. */
    public static final String WORKSHEET_WRITER         = "51F952E9-A914-4F12-B1CC-2F6C1F3637D7";

    /** Replacing plugin: writes the styles XML part ({@code styles.xml}). */
    public static final String STYLE_WRITER             = "009D7028-E8D9-4BB6-B5C7-F6D5EA2BA01F";

    /** Replacing plugin: writes the shared-strings XML part ({@code sharedStrings.xml}). */
    public static final String SHARED_STRINGS_WRITER    = "731BF436-E28D-4136-BEF4-394D2CC65E01";

    /** Replacing plugin: writes the application metadata part ({@code app.xml}). */
    public static final String META_DATA_APP_WRITER     = "49910428-CACB-475A-B39D-833D384DADE8";

    /** Replacing plugin: writes the core metadata part ({@code core.xml}). */
    public static final String META_DATA_CORE_WRITER    = "19C28EEF-D80E-4A22-9B30-26376C7512FE";

    /** Replacing plugin: writes the theme XML part ({@code theme/theme1.xml}). */
    public static final String THEME_WRITER             = "62E3A926-08F3-4343-ACCE-2A42096C3235";

    /** Replacing plugin: writes color information within the theme. */
    public static final String COLOR_WRITER             = "7276A073-55D2-482A-B5CD-AB752A70EA9D";

    // -------------------------------------------------------------------------
    // GENERAL WRITER QUEUES (general — execute around the full write pass)
    // -------------------------------------------------------------------------

    /** Queue plugin: executed <em>before</em> the standard write pass begins. */
    public static final String WRITER_PREPEND_QUEUE          = "772C4BF6-ED81-4127-80C7-C99D2B5C5EEC";

    /** Queue plugin: used to register additional OOXML package parts during writing. */
    public static final String WRITER_PACKAGE_REGISTRY_QUEUE = "C0CE40AC-14D5-4403-A5A3-018C6057A80E";

    /** Queue plugin: executed <em>after</em> the standard write pass completes. */
    public static final String WRITER_APPEND_QUEUE           = "04F73656-C355-40A9-9E68-CB21329F3E53";

    // -------------------------------------------------------------------------
    // INLINE QUEUE WRITERS (execute immediately after the named component)
    // -------------------------------------------------------------------------

    /** Queue plugin: executed immediately after the workbook writer component. */
    public static final String WORKBOOK_INLINE_WRITER        = "E69CEC04-A5CD-4DC2-9517-88F895C5CB1E";

    /** Queue plugin: executed immediately after the worksheet writer component. */
    public static final String WORKSHEET_INLINE_WRITER       = "E0F6C065-00F8-4A67-AFAF-F358342845BC";

    /** Queue plugin: executed immediately after the style writer component. */
    public static final String STYLE_INLINE_WRITER           = "E9358F10-DD9B-4C5B-9BBB-DC32D5EB0DBB";

    /** Queue plugin: executed immediately after the shared-strings writer component. */
    public static final String SHARED_STRINGS_INLINE_WRITER  = "1E87131E-E6BA-4292-B4E5-55B73233D3F5";

    /** Queue plugin: executed immediately after the application-metadata writer component. */
    public static final String META_DATA_APP_INLINE_WRITER   = "AB45D7E1-7FF9-43D9-B482-91D677A7D614";

    /** Queue plugin: executed immediately after the core-metadata writer component. */
    public static final String META_DATA_CORE_INLINE_WRITER  = "85AC02E3-1F92-4921-BC69-39B3F328ABCD";

    /** Queue plugin: executed immediately after the theme writer component. */
    public static final String THEME_INLINE_WRITER           = "4CB6FD0E-AB69-40E9-B048-06B0E00C892D";

    // -------------------------------------------------------------------------
    // READERS
    // -------------------------------------------------------------------------

    /** Replacing plugin: reads the password / encryption information. */
    public static final String PASSWORD_READER          = "1090EEDC-27AB-4A90-AAAB-E9B02C086082";

    /** Replacing plugin: reads the workbook XML part ({@code workbook.xml}). */
    public static final String WORKBOOK_READER          = "B8C3405A-081C-453B-9C88-6A4BD7F5359B";

    /** Replacing plugin: reads individual worksheet XML parts. */
    public static final String WORKSHEET_READER         = "1DE75D75-5BF9-48EA-9387-DCF5459EC401";

    /** Replacing plugin: reads the styles XML part ({@code styles.xml}). */
    public static final String STYLE_READER             = "67AAB19A-4BF1-41B4-BC86-8C5BB5BB91F6";

    /** Replacing plugin: reads the shared-strings XML part ({@code sharedStrings.xml}). */
    public static final String SHARED_STRINGS_READER    = "FF9BC0E6-59BF-4A16-B289-3F2AFD568438";

    /** Replacing plugin: reads the application metadata part ({@code app.xml}). */
    public static final String META_DATA_APP_READER     = "28C59145-7BB8-416F-BAC9-0130DD8557F9";

    /** Replacing plugin: reads the core metadata part ({@code core.xml}). */
    public static final String META_DATA_CORE_READER    = "B53F0F3E-71FF-43F0-B60C-C3478DE65788";

    /** Replacing plugin: reads the theme XML part ({@code theme/theme1.xml}). */
    public static final String THEME_READER             = "B4733D00-B596-4440-8E33-A803289848BC";

    /** Replacing plugin: reads the relationship XML parts ({@code .rels} files). */
    public static final String RELATIONSHIP_READER      = "DB9AF89B-6181-4F94-A666-5AB70840EDDF";

    // -------------------------------------------------------------------------
    // GENERAL READER QUEUES (general — execute around the full read pass)
    // -------------------------------------------------------------------------

    /** Queue plugin: executed <em>before</em> the standard read pass begins. */
    public static final String READER_PREPEND_QUEUE          = "658A903B-512D-490C-A99B-40C0B0947CBF";

    /** Queue plugin: used to register additional OOXML package parts during reading. */
    public static final String READER_PACKAGE_REGISTRY_QUEUE = "1DD50B15-6EB8-451B-A6A8-C9265A8EF55C";

    /** Queue plugin: executed <em>after</em> the standard read pass completes. */
    public static final String READER_APPEND_QUEUE           = "69EE822E-910E-4E6B-BC5B-8F27629933AF";

    // -------------------------------------------------------------------------
    // READER INLINE QUEUES (execute immediately after the named component)
    // -------------------------------------------------------------------------

    /** Queue plugin: executed immediately after the workbook reader component. */
    public static final String WORKBOOK_INLINE_READER        = "33782BED-FCBA-4BE1-911A-5327C64B9580";

    /** Queue plugin: executed immediately after the worksheet reader component. */
    public static final String WORKSHEET_INLINE_READER       = "20BE8320-9B90-41D2-8580-E1FE05DDC881";

    /** Queue plugin: executed immediately after the style reader component. */
    public static final String STYLE_INLINE_READER           = "9AC00387-E677-4F1C-88D6-558DAE6FF764";

    /** Queue plugin: executed immediately after the shared-strings reader component. */
    public static final String SHARED_STRINGS_INLINE_READER  = "3730F89E-CD7C-4BD8-B6AC-A18D803ADB2B";

    /** Queue plugin: executed immediately after the application-metadata reader component. */
    public static final String META_DATA_APP_INLINE_READER   = "789AFD19-31C5-409A-86C6-7CF5CC49B9C1";

    /** Queue plugin: executed immediately after the core-metadata reader component. */
    public static final String META_DATA_CORE_INLINE_READER  = "64A26388-EAD1-4435-AC07-A7FF18DCEEB7";

    /** Queue plugin: executed immediately after the theme reader component. */
    public static final String THEME_INLINE_READER           = "4B44E8A8-4560-44EB-8B24-5E11FDC04971";

    /** Queue plugin: executed immediately after the relationship reader component. */
    public static final String RELATIONSHIP_INLINE_READER    = "E474D078-FBBC-49BE-B0B8-6086C07023DA";

    // -------------------------------------------------------------------------
    // ENTITIES
    // -------------------------------------------------------------------------

    /** Replacing plugin: provides worksheet definition entities during workbook processing. */
    public static final String WORKSHEET_DEFINITION_ENTITY  = "40CF0799-E4E7-4EA7-925F-BB6C9E8F588A";

    /** Replacing plugin: provides selected-worksheet entities during workbook processing. */
    public static final String SELECTED_WORKSHEET_ENTITY    = "DD9B5E9B-2276-484D-B36A-B1F5EB6EE08A";

    /** Replacing plugin: provides relationship entities during workbook processing. */
    public static final String RELATIONSHIP_ENTITY          = "F2DECC2C-544A-4B22-8C6E-386464586E60";

    /** Replacing plugin: provides style entities during workbook processing. */
    public static final String STYLE_ENTITY                 = "638F9F5A-334A-49A1-BE07-1F2F3BFB70C4";
}
