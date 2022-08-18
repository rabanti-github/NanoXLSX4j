package ch.rabanti.nanoxlsx4j.reader;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.InputStream;

import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.Cell;
import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.styles.Border;

public class ReadFallbacksTest {
    

    @DisplayName("Test of the fallback behavior on unexpected border types")
    @Test()
    void readUnknownBorderTypeTest() throws Exception {
        // Cell A1 contains a border style with unknown line type
        // This causes neither in Excel a crash, nor should the library crash
        Cell cell = getCell("unknown_style_enums.xlsx");
        assertEquals(Border.StyleValue.none, cell.getCellStyle().getBorder().getTopStyle());
        assertEquals(Border.StyleValue.none, cell.getCellStyle().getBorder().getBottomStyle());
        assertEquals(Border.StyleValue.none, cell.getCellStyle().getBorder().getLeftStyle());
        assertEquals(Border.StyleValue.none, cell.getCellStyle().getBorder().getRightStyle());
        assertEquals(Border.StyleValue.none, cell.getCellStyle().getBorder().getDiagonalStyle());
    }

    @DisplayName("Test of the fallback behavior on unexpected pattern fill types")
    @Test()
    void readUnknownPatternFillTypeTest() throws Exception {
        // The file contains a pattern fill definition with an unknown value
        // This causes neither in Excel a crash, nor should the library crash
        Cell cell = getCell("unknown_style_enums.xlsx");
        assertEquals(Fill.PatternValue.none, cell.getCellStyle().getFill().getPatternFill());
    }

    @DisplayName("Test of the fallback behavior on unexpected vertical align font types")
    @Test()
    void readUnknownFontVerticalAlignTypeTest() throws Exception {
        // The file contains a font definition with an unknown vertical align value
        // This causes an auto-fixing action in Excel (but not a crash). The library will auto-fix this too
        Cell cell = getCell("unknown_style_enums.xlsx");
        assertEquals(Font.VerticalAlignValue.none, cell.getCellStyle().getFont().getVerticalAlign());
    }

    @DisplayName("Test of the fallback behavior on unexpected horizontal align cellXF types")
    @Test()
    void readUnknownCellXfHorizontalAlignTypeTest() throws Exception {
        // The file contains a CellXF definition with an unknown horizontal align value
        // This causes neither in Excel a crash, nor should the library crash
        Cell cell = getCell("unknown_style_enums.xlsx");
        assertEquals(CellXf.HorizontalAlignValue.none, cell.getCellStyle().getCellXf().getHorizontalAlign());
    }

    @DisplayName("Test of the fallback behavior on unexpected vertical align cellXF types")
    @Test()
    void readUnknownCellXfVerticalAlignTypeTest() throws Exception {
        // The file contains a CellXF definition with an unknown vertical align value
        // This causes neither in Excel a crash, nor should the library crash
        Cell cell = getCell("unknown_style_enums.xlsx");
        assertEquals(CellXf.VerticalAlignValue.none, cell.getCellStyle().getCellXf().getVerticalAlign());
    }

    private static Cell getCell(String resourceName) throws IOException, java.io.IOException {
        InputStream stream = TestUtils.getResource(resourceName);
        Workbook workbook = Workbook.load(stream);
        Cell cell = workbook.getWorksheets().get(0).getCell("A1");
        return cell;
    }

}
