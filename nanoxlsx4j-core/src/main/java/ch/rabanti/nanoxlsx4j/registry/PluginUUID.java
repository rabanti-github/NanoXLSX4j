/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.registry;

/**
 * UUIDs used to register NanoXLSX plug-ins and identify plug-in-owned data entities.
 * <p>
 * Existing values are compatibility identifiers shared with NanoXLSX .NET and must never be changed.
 */
public final class PluginUUID {
// -------------------------------
// Writer UUIDs
    /** UUID for the password writer when a workbook is saved. */
    public static final String PASSWORD_WRITER = "8106E566-60D6-45DB-BF87-33AB3882C019";
    /** UUID for the workbook writer when a workbook is saved. */
    public static final String WORKBOOK_WRITER = "D4272E3A-AC56-4524-9B9F-7B1448DF536B";
    /** UUID for the worksheet writer when a workbook is saved. */
    public static final String WORKSHEET_WRITER = "51F952E9-A914-4F12-B1CC-2F6C1F3637D7";
    /** UUID for the style writer when a workbook is saved. */
    public static final String STYLE_WRITER = "009D7028-E8D9-4BB6-B5C7-F6D5EA2BA01F";
    /** UUID for the shared strings writer when a workbook is saved. */
    public static final String SHARED_STRINGS_WRITER = "731BF436-E28D-4136-BEF4-394D2CC65E01";
    /** UUID for the application metadata writer when a workbook is saved. */
    public static final String METADATA_APP_WRITER = "49910428-CACB-475A-B39D-833D384DADE8";
    /** UUID for the core metadata writer when a workbook is saved. */
    public static final String METADATA_CORE_WRITER = "19C28EEF-D80E-4A22-9B30-26376C7512FE";
    /** UUID for the theme writer when a workbook is saved. */
    public static final String THEME_WRITER = "62E3A926-08F3-4343-ACCE-2A42096C3235";
    /** UUID for the color writer that provides general color-definition writing functionality. */
    public static final String COLOR_WRITER = "7276A073-55D2-482A-B5CD-AB752A70EA9D";
    /** UUID for the initial processor that performs checks and actions before any writer plug-in is executed. */
    public static final String PREPARING_PROCESSOR = "8156D4A2-BADB-4196-B5FC-4BE1CC743CFE";
// -------------------------------
// General writer UUIDs
    /** UUID for plug-ins executed before the regular XLSX writers. */
    public static final String WRITER_PREPENDING_QUEUE = "772C4BF6-ED81-4127-80C7-C99D2B5C5EEC";
    /** UUID for plug-ins that register additional package parts for the XLSX building process. */
    public static final String WRITER_PACKAGE_REGISTRY_QUEUE = "C0CE40AC-14D5-4403-A5A3-018C6057A80E";
    /** UUID for plug-ins executed after the regular XLSX writers. */
    public static final String WRITER_APPENDING_QUEUE = "04F73656-C355-40A9-9E68-CB21329F3E53";
// -------------------------------
// Inline queue writer UUIDs
    /** UUID for queued writers executed immediately after the workbook writer. */
    public static final String WORKBOOK_INLINE_WRITER = "E69CEC04-A5CD-4DC2-9517-88F895C5CB1E";
    /** UUID for queued writers executed immediately after the worksheet writer. */
    public static final String WORKSHEET_INLINE_WRITER = "E0F6C065-00F8-4A67-AFAF-F358342845BC";
    /** UUID for queued writers executed immediately after the style writer. */
    public static final String STYLE_INLINE_WRITER = "E9358F10-DD9B-4C5B-9BBB-DC32D5EB0DBB";
    /** UUID for queued writers executed immediately after the shared strings writer. */
    public static final String SHARED_STRINGS_INLINE_WRITER = "1E87131E-E6BA-4292-B4E5-55B73233D3F5";
    /** UUID for queued writers executed immediately after the application metadata writer. */
    public static final String METADATA_APP_INLINE_WRITER = "AB45D7E1-7FF9-43D9-B482-91D677A7D614";
    /** UUID for queued writers executed immediately after the core metadata writer. */
    public static final String METADATA_CORE_INLINE_WRITER = "85AC02E3-1F92-4921-BC69-39B3F328ABCD";
    /** UUID for queued writers executed immediately after the theme writer. */
    public static final String THEME_INLINE_WRITER = "4CB6FD0E-AB69-40E9-B048-06B0E00C892D";
    /** UUID for queued processors executed immediately after the initial preparing processor. */
    public static final String PREPARING_INLINE_PROCESSOR = "6CB92B28-25A5-4E80-92BA-5F7D904E993C";
    /**
     * UUID for queued processors executed immediately before the core compatibility check.
     * <p>
     * Unlike other inline plug-ins, these plug-ins run before the core compatibility processor performs its checks,
     * allowing otherwise disabled features to be enabled. The core compatibility processor cannot be replaced.
     */
    public static final String COMPATIBILITY_INLINE_PROCESSOR = "B86C2CD1-1E94-465A-AEA1-BF8C4108533D";
// -------------------------------
// Writer entity UUIDs
    /** UUID for the highest currently defined package-part order number when a workbook is saved. */
    public static final String LAST_PACKAGE_ORDER_NUMBER = "C6165F97-98A7-478C-AD88-D21FDC58D234";
    /** UUID for workbook relationship IDs assigned to registered package parts when a workbook is saved. */
    public static final String PACKAGE_PART_RELATIONSHIP_ID = "66C4B211-A4F6-404A-974D-4B3E260F8E29";
// -------------------------------
// Reader UUIDs
    /** UUID for the password reader when a workbook is loaded. */
    public static final String PASSWORD_READER = "1090EEDC-27AB-4A90-AAAB-E9B02C086082";
    /** UUID for the workbook reader when a workbook is loaded. */
    public static final String WORKBOOK_READER = "B8C3405A-081C-453B-9C88-6A4BD7F5359B";
    /** UUID for the worksheet reader when a workbook is loaded. */
    public static final String WORKSHEET_READER = "1DE75D75-5BF9-48EA-9387-DCF5459EC401";
    /** UUID for the style reader when a workbook is loaded. */
    public static final String STYLE_READER = "67AAB19A-4BF1-41B4-BC86-8C5BB5BB91F6";
    /** UUID for the shared strings reader when a workbook is loaded. */
    public static final String SHARED_STRINGS_READER = "FF9BC0E6-59BF-4A16-B289-3F2AFD568438";
    /** UUID for the application metadata reader when a workbook is loaded. */
    public static final String METADATA_APP_READER = "28C59145-7BB8-416F-BAC9-0130DD8557F9";
    /** UUID for the core metadata reader when a workbook is loaded. */
    public static final String METADATA_CORE_READER = "B53F0F3E-71FF-43F0-B60C-C3478DE65788";
    /** UUID for the theme reader when a workbook is loaded. */
    public static final String THEME_READER = "B4733D00-B596-4440-8E33-A803289848BC";
    /** UUID for the legacy relationship reader when a workbook is loaded. */
    public static final String RELATIONSHIP_READER = "DB9AF89B-6181-4F94-A666-5AB70840EDDF";
    /** UUID for the package relationship discovery reader when a workbook is loaded. */
    public static final String DISCOVERY_READER = "08FE5C9F-C8A3-4B46-A74F-D2BF3F083CC4";
    /**
     * UUID for the processor that performs finalizing tasks after workbook parts and appended queue readers run.
     * Such plug-ins do not process a stream.
     */
    public static final String FINALIZING_PROCESSOR = "7B6B70B7-F2E1-433A-8171-018E93DE43A0";
// -------------------------------
// General reader queue UUIDs
    /** UUID for plug-ins executed before the regular XLSX readers. */
    public static final String READER_PREPENDING_QUEUE = "658A903B-512D-490C-A99B-40C0B0947CBF";
    /** UUID for plug-ins that register additional package parts for the XLSX reading process. */
    public static final String READER_PACKAGE_REGISTRY_QUEUE = "1DD50B15-6EB8-451B-A6A8-C9265A8EF55C";
    /** UUID for plug-ins executed after the regular XLSX readers. */
    public static final String READER_APPENDING_QUEUE = "69EE822E-910E-4E6B-BC5B-8F27629933AF";
// -------------------------------
// Inline queue reader UUIDs
    /** UUID for queued readers executed immediately after the workbook reader. */
    public static final String WORKBOOK_INLINE_READER = "33782BED-FCBA-4BE1-911A-5327C64B9580";
    /** UUID for a queued reader executed immediately after the worksheet reader. */
    public static final String WORKSHEET_INLINE_READER = "20BE8320-9B90-41D2-8580-E1FE05DDC881";
    /** UUID for a queued reader executed immediately after the style reader. */
    public static final String STYLE_INLINE_READER = "9AC00387-E677-4F1C-88D6-558DAE6FF764";
    /** UUID for a queued reader executed immediately after the shared strings reader. */
    public static final String SHARED_STRINGS_INLINE_READER = "3730F89E-CD7C-4BD8-B6AC-A18D803ADB2B";
    /** UUID for a queued reader executed immediately after the application metadata reader. */
    public static final String METADATA_APP_INLINE_READER = "789AFD19-31C5-409A-86C6-7CF5CC49B9C1";
    /** UUID for a queued reader executed immediately after the core metadata reader. */
    public static final String METADATA_CORE_INLINE_READER = "64A26388-EAD1-4435-AC07-A7FF18DCEEB7";
    /** UUID for a queued reader executed immediately after the theme reader. */
    public static final String THEME_INLINE_READER = "4B44E8A8-4560-44EB-8B24-5E11FDC04971";
    /** UUID for a queued reader executed immediately after the relationship reader. */
    public static final String RELATIONSHIP_INLINE_READER = "E474D078-FBBC-49BE-B0B8-6086C07023DA";
    /**
     * UUID for a queued processor executed immediately after the finalizing reader.
     * Such processors do not process a stream.
     */
    public static final String FINALIZING_INLINE_PROCESSOR = "0DAF416F-C735-4EB1-962A-859495E034DC";
// -------------------------------
// Reader entity UUIDs
    /** UUID for the worksheet definitions section when a workbook is read. */
    public static final String WORKSHEET_DEFINITION_ENTITY = "40CF0799-E4E7-4EA7-925F-BB6C9E8F588A";
    /** UUID for the selected worksheet when a workbook is read. */
    public static final String SELECTED_WORKSHEET_ENTITY = "DD9B5E9B-2276-484D-B36A-B1F5EB6EE08A";
    /** UUID for the worksheet relationship section when a workbook is read. */
    public static final String RELATIONSHIP_ENTITY = "F2DECC2C-544A-4B22-8C6E-386464586E60";
    /** UUID for styles when a workbook is read. */
    public static final String STYLE_ENTITY = "638F9F5A-334A-49A1-BE07-1F2F3BFB70C4";
    /** UUID for the defined names section when a workbook is read. */
    public static final String DEFINED_NAME_ENTITY = "7774EFED-65A6-4AD7-9870-20DE4D09DB41";
    /** UUID for the relationship discovery catalog when a workbook is read. */
    public static final String DISCOVERY_CATALOG_ENTITY = "A3FB109E-F7F8-4FA1-9AAF-CF3C5A62A9A8";
// -------------------------------
// Feature reader UUIDs
    /** UUID identifying the writer feature for external links such as external workbooks. */
    public static final String WRITE_EXTERNAL_LINK_FEATURE = "D25A7462-CC8B-491E-B620-4A8EF81DE717";
    // TODO add here new features (keep in sync with region "Feature Writer UUIDs"; use different UUIDs)
// -------------------------------
// Feature writer UUIDs
    /** UUID identifying the reader feature for external links such as external workbooks. */
    public static final String READ_EXTERNAL_LINK_FEATURE = "8E62F688-B7C6-4464-AF2A-BDBFF5F04AB5";
    // TODO add here new features (keep in sync with region "Feature Reader UUIDs"; use different UUIDs)

    private PluginUUID() {
        // Do not instantiate
    }
}
