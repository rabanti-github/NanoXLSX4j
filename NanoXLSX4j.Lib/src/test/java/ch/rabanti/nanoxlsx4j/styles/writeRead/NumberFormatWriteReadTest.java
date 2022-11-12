package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.Helper;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.io.ByteArrayOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class NumberFormatWriteReadTest {

    @DisplayName("Test of the 'customFormatID' value when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with value '{1}' should lead to a cell with the same format number")
    @CsvSource(
            {
                    "164, test",
                    "165, 0.5f",
                    "200, 22",
                    "2000, true",}
    )
    void customFormatIDFormatTest(int styleValue, Object value) {
        Style style = new Style();
        style.getNumberFormat().setCustomFormatID(styleValue);
        style.getNumberFormat().setNumber(NumberFormat.FormatNumber.custom); // Mandatory
        style.getNumberFormat().setCustomFormatCode("#.##"); // Mandatory
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getCustomFormatID());
    }

    @DisplayName("Test of the failing save attempt of 'customFormatID' value when writing and reading a NumberFormat style with missing CustomFormatCode")
    @ParameterizedTest(name = "Given format number {0} with value '{1}' should lead to an exception")
    @CsvSource(
            {
                    "164, test",
                    "165, 0.5f",
                    "200, 22",
                    "2000, true",}
    )
    void customFormatIDFormatFailTest(int styleValue, Object value) {
        Workbook workbook = new Workbook("worksheet1");
        Style style = new Style();
        style.getNumberFormat().setCustomFormatID(styleValue);
        style.getNumberFormat().setNumber(NumberFormat.FormatNumber.custom); // Mandatory
        workbook.getCurrentWorksheet().addNextCell("test", style);
        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        assertThrows(Exception.class, () -> workbook.saveAsStream(stream));
    }

    @DisplayName("Test of the 'customFormatCode' value when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with value '{1}' should lead to a cell with the same format number")
    @CsvSource(
            {
                    "#, test",
                    "'\\', 0.5f",
                    "'\\\\', ''",
                    "' \\.\\ ', false",
                    "' ', 22",
                    "ABCDE, true",}
    )
    void customFormatCodeNumberFormatTest(String styleValue, Object value) {
        Style style = new Style();
        style.getNumberFormat().setCustomFormatCode(styleValue);
        style.getNumberFormat().setNumber(NumberFormat.FormatNumber.custom); // Mandatory
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getCustomFormatCode());
        assertEquals(NumberFormat.FormatNumber.custom, cell.getCellStyle().getNumberFormat().getNumber());
        assertTrue(cell.getCellStyle().getNumberFormat().isCustomFormat());
    }

    @DisplayName("Test of the 'formatNumber' value when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with value '{1}' should lead to a cell with the same format number")
    @CsvSource(
            {
                    "format_1, test",
                    "format_2, 0.5f",
                    "format_3, 22",
                    "format_4, true",
                    "format_5, ",
                    "format_6, -1",
                    "format_7, -22.222f",
                    "format_8, false",
                    "format_9, 0",
                    "format_10, Æ",
                    "format_11, test",
                    "format_12, 0.5f",
                    "format_13, 22",
                    "format_14, true",
                    "format_15, ''",
                    "format_16, -1",
                    "format_17, -22.222f",
                    "format_18, false",
                    "format_19, noDate",
                    "format_20, Æ",
                    "format_21, test",
                    "format_22, noDate",
                    "format_37, 22",
                    "format_38, true",
                    "format_39, ",
                    "format_40, -1",
                    "format_45, -22.222f",
                    "format_46, false",
                    "format_47, noDate",
                    "format_48, Æ",
                    "format_49, test",}
    )
    void numberNumberFormatTest(NumberFormat.FormatNumber styleValue, Object value) {
        Style style = new Style();
        style.getNumberFormat().setNumber(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getNumber());
    }

    @DisplayName("Test of the 'formatNumber' value with date formats when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with given value '{1}' should lead to a cell with the same format number and the expected value '{2}'")
    @CsvSource(
            {
                    "format_14, 1000, 26.09.1902",
                    "format_15, 1000, 26.09.1902",
                    "format_16, 1000, 26.09.1902",
                    "format_17, 1000, 26.09.1902",
                    "format_22, 1000, 26.09.1902",}
    )
    void numberFormatTest2(NumberFormat.FormatNumber styleValue, int value, String expected) throws ParseException {
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy", Locale.US);
        Date expectedValue = formatter.parse(expected);
        Style style = new Style();
        style.getNumberFormat().setNumber(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, expectedValue, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getNumber());
        assertEquals(expectedValue, cell.getValue());
    }

    @DisplayName("Test of the 'formatNumber' value with time formats when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with given value '{1}' should lead to a cell with the same format number and the expected value '{2}'")
    @CsvSource(
            {
                    "format_19, 0.5, 12:00:00",
                    "format_20, 0.5, 12:00:00",
                    "format_21, 0.5, 12:00:00",
                    "format_45, 0.5, 12:00:00",
                    "format_46, 0.5, 12:00:00",
                    "format_47, 0.5, 12:00:00",}
    )
    void numberFormatTest3(NumberFormat.FormatNumber styleValue, float value, String expected) {
        Duration expectedValue = Helper.parseTime(expected, "HH:mm:ss", Locale.US);
        Style style = new Style();
        style.getNumberFormat().setNumber(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, expectedValue, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getNumber());
        assertEquals(expectedValue, cell.getValue());
    }

}
