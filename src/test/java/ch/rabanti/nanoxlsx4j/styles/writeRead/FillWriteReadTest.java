package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Address;
import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

public class FillWriteReadTest {

    @DisplayName("Test of the writing and reading of the foreground color style value")
    @ParameterizedTest(name = "Given foreground color {0} with value {1} should lead to a cell with this color")
    @CsvSource({
            "FFAACC00, 'test'",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void foregroundColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getFill().setForegroundColor(color);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getFill().getForegroundColor());
        assertNotEquals(Fill.PatternValue.none, cell.getCellStyle().getFill().getPatternFill());
    }

    @DisplayName("Test of the writing and reading of the background color style value")
    @ParameterizedTest(name = "Given background color {0} with value {1} should lead to a cell with this color")
    @CsvSource({
            "FFAACC00, 'test'",
            "FFAADD00, 0.5f",
            "FFDDCC00, true",
            "FFAACCDD, null",
    })
    void backgroundColorTest(String color, Object value)
    {
        Style style = new Style();
        style.getFill().setBackgroundColor(color);
        style.getFill().setPatternFill(Fill.PatternValue.darkGray);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(color, cell.getCellStyle().getFill().getBackgroundColor());
        assertEquals(Fill.PatternValue.darkGray, cell.getCellStyle().getFill().getPatternFill());
    }

    @DisplayName("Test of the writing and reading of the pattern value for fills styles")
    @ParameterizedTest(name = "Given pattern value {0} with value {1} should lead to a cell with this pattern")
    @CsvSource({
            "solid, 'test'",
            "darkGray, 0.5f",
            "gray0625, true",
            "gray125, null",
            "lightGray, ''",
            "mediumGray, 0",
            "none, true",
    })
    void patternValueTest(Fill.PatternValue pattern, Object value)
    {
        Style style = new Style();
        style.getFill().setPatternFill(pattern);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");

        assertEquals(pattern, cell.getCellStyle().getFill().getPatternFill());
    }



}
