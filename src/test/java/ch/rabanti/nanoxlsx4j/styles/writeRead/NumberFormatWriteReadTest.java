package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class NumberFormatWriteReadTest {
    @DisplayName("Test of the 'formatNumber' value when writing and reading a NumberFormat style")
    @ParameterizedTest(name = "Given format number {0} with value '{1}' should lead to a cell with the same format number")
    @CsvSource({
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
            "format_49, test",
            "custom, 0.5f",
    })
    void numberNumberFormatTest(NumberFormat.FormatNumber styleValue, Object value)
    {
        Style style = new Style();
        style.getNumberFormat().setNumber(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getNumberFormat().getNumber());
    }

}
