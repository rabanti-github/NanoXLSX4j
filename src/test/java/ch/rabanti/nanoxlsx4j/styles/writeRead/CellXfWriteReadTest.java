package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class CellXfWriteReadTest {

    @DisplayName("Test of the 'ForceApplyAlignment' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given ForceApplyAlignment {0} with value '{1}' should lead to a cell with the same ForceApplyAlignment value")
    @CsvSource({
            "true, test",
            "false, 0.5f",
    })
    void forceApplyAlignmentCellXfTest(boolean styleValue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setForceApplyAlignment(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getCellXf().isForceApplyAlignment());
    }

    @DisplayName("Test of the 'Hidden' and 'Locked' values when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given Hidden {0} and Locked {1} with value '{2}' should lead to a cell with the same Hidden/Locked values")
    @CsvSource({
            "false, false, 'test'",
            "false, true, 0.5f",
            "true, false, 22",
            "true, true, true",
    })
    void hiddenCellXfTest(boolean hiddenStyleValue, boolean lockedStylevalue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setHidden(hiddenStyleValue);
        style.getCellXf().setLocked(lockedStylevalue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(hiddenStyleValue, cell.getCellStyle().getCellXf().isHidden());
        assertEquals(lockedStylevalue, cell.getCellStyle().getCellXf().isLocked());
    }

    @DisplayName("Test of the 'Alignment' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given Alignment {0} with value '{1}' should lead to a cell with the same Alignment value")
    @CsvSource({
            "shrinkToFit, test",
            "wrapText, 0.5f",
            "none, true",
    })
    void alignmentCellXfTest(CellXf.TextBreakValue styleValue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setAlignment(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getCellXf().getAlignment());
    }

    @DisplayName("Test of the 'HorizontalAlign' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given HorizontalAlign {0} with value '{1}' should lead to a cell with the same HorizontalAlign value")
    @CsvSource({
            "justify, test",
            "center, 0.5f",
            "centerContinuous, true",
            "distributed, 22",
            "fill, false",
            "general, ",
            "left, -2.11f",
            "right, test",
            "none,  ",
    })
    void horizontalAlignCellXfTest(CellXf.HorizontalAlignValue styleValue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setHorizontalAlign(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getCellXf().getHorizontalAlign());
    }

    @DisplayName("Test of the 'VerticalAlign' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given VerticalAlign {0} with value '{1}' should lead to a cell with the same VerticalAlign value")
    @CsvSource({
            "justify, test",
            "center, 0.5f",
            "bottom, true",
            "top, 22",
            "distributed, false",
            "none, ' '",
    })
    void verticalAlignCellXfTest(CellXf.VerticalAlignValue styleValue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setVerticalAlign(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getCellXf().getVerticalAlign());
    }

    @DisplayName("Test of the 'Indent' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given indent {0} with horizontal alignment {1} and value {3} should lead to a cell with the expected indent {2}")
    @CsvSource({
            "0, left, 0, test",
            "0, right, 0, test",
            "0, distributed, 0, test",
            "0, center, 0, test",
            "1, left, 1, 0.5f",
            "1, right, 1, 0.5f",
            "1, distributed, 1, 0.5f",
            "1, center, 0, 0.5f",
            "5, left, 5, true",
            "5, right, 5, true",
            "5, distributed, 5, true",
            "5, center, 0, true",
            "64, left, 64, 22",
            "64, right, 64, 22",
            "64, distributed, 64, 22",
            "64, center, 0, 22",
    })
    void indentCellXfTest(int styleValue, CellXf.HorizontalAlignValue alignValue, int expectedIndent, Object value)
    {
        Style style = new Style();
        style.getCellXf().setHorizontalAlign(alignValue);
        style.getCellXf().setIndent(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(expectedIndent, cell.getCellStyle().getCellXf().getIndent());
        assertEquals(alignValue, cell.getCellStyle().getCellXf().getHorizontalAlign()
        );
    }

    @DisplayName("Test of the 'TextRotation' value when writing and reading a CellXF style")
    @ParameterizedTest(name = "Given TextRotation {0} with value '{1}' should lead to a cell with the same TextRotation value")
    @CsvSource({
            "0, test",
            "1, 0.5f",
            "-1, true",
            "45,22",
            "-45, -0.11f",
            "90, ",
            "-90,  ",
    })
    void textRotationCellXfTest(int styleValue, Object value)
    {
        Style style = new Style();
        style.getCellXf().setTextRotation(styleValue);
        Cell cell = TestUtils.saveAndReadStyledCell(value, style, "A1");
        assertEquals(styleValue, cell.getCellStyle().getCellXf().getTextRotation());
    }




}
