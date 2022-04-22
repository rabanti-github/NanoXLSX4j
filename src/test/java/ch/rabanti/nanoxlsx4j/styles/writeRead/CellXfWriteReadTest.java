package ch.rabanti.nanoxlsx4j.styles.writeRead;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
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
    @ParameterizedTest(name = "Given Hidden {0} and Locked {1} with value '{2}' should lead to a cell with the same Hidden value")
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

}
