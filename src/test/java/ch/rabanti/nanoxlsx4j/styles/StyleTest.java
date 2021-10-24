package ch.rabanti.nanoxlsx4j.styles;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.*;

public class StyleTest {

    @DisplayName("Test of the get and set function of the Border field")
    @Test()
    void currentBorderTest() {
        Style style = new Style();
        Border border = new Border();
        assertNotNull(style.getBorder());
        assertEquals(border.hashCode(), style.getBorder().hashCode());
        style.setBorder(border);
        border.setBottomColor("FFAABBCC");
        assertEquals("FFAABBCC", style.getBorder().getBottomColor());
    }


    @DisplayName("Test of the get and set function of the cellXf field")
    @Test()
    void currentCellXfTest() {
        Style style = new Style();
        CellXf cellXf = new CellXf();
        assertNotNull(style.getCellXf());
        assertEquals(cellXf.hashCode(), style.getCellXf().hashCode());
        style.setCellXf(cellXf);
        cellXf.setIndent(5);
        assertEquals(5, style.getCellXf().getIndent());
    }


    @DisplayName("Test of the get and set function of the fill field")
    @Test()
    void currentFillTest() {
        Style style = new Style();
        Fill fill = new Fill();
        assertNotNull(style.getFill());
        assertEquals(fill.hashCode(), style.getFill().hashCode());
        style.setFill(fill);
        fill.setBackgroundColor("AACCBBDD");
        assertEquals("AACCBBDD", style.getFill().getBackgroundColor());
    }


    @DisplayName("Test of the get and set function of the font field")
    @Test()
    void currentFontTest() {
        Style style = new Style();
        Font font = new Font();
        assertNotNull(style.getFont());
        assertEquals(font.hashCode(), style.getFont().hashCode());
        style.setFont(font);
        font.setName("Sans Serif");
        assertEquals("Sans Serif", style.getFont().getName());
    }


    @DisplayName("Test of the get and set function of the numberFormat field")
    @Test()
    void currentNumberFormatTest() {
        Style style = new Style();
        NumberFormat numberFormat = new NumberFormat();
        assertNotNull(style.getFill());
        assertEquals(numberFormat.hashCode(), style.getNumberFormat().hashCode());
        style.setNumberFormat(numberFormat);
        numberFormat.setNumber(NumberFormat.FormatNumber.format_15);
        assertEquals(NumberFormat.FormatNumber.format_15, style.getNumberFormat().getNumber());
    }


    @DisplayName("Test of the get and set function of the name field")
    @Test()
    void nameTest() {
        Style style = new Style();
        assertEquals(Integer.toString(style.hashCode()), style.getName());
        style.setName("Test");
        assertEquals("Test", style.getName());
    }


    @DisplayName("Test of the get function of the ssInternalStyle field")
    @Test()
    void isInternalStyleTest() {
        Style style = new Style();
        assertFalse(style.isInternalStyle());
        Style internalStyle = new Style("test", 0, true);
        assertTrue(internalStyle.isInternalStyle());
    }


    @DisplayName("Test of the get and set function of the internalID field")
    @Test()
    void internalIDTest() {
        Style style = new Style();
        assertNull(style.getInternalID());
        style.setInternalID(962);
        assertEquals(962, style.getInternalID());
    }


    @DisplayName("Test of the default constructor")
    @Test()
    void constructorTest() {
        Style style = new Style();
        assertNotNull(style.getBorder());
        assertNotNull(style.getCellXf());
        assertNotNull(style.getFill());
        assertNotNull(style.getFont());
        assertNotNull(style.getNumberFormat());
        assertNotNull(style.getName());
        assertNull(style.getInternalID());
    }


    @DisplayName("Test of the constructor with a name")
    @Test()
    void constructorTest2() {
        Style style = new Style("test1");
        assertNotNull(style.getBorder());
        assertNotNull(style.getCellXf());
        assertNotNull(style.getFill());
        assertNotNull(style.getFont());
        assertNotNull(style.getNumberFormat());
        assertEquals("test1", style.getName());
        assertNull(style.getInternalID());
    }


    @DisplayName("Test of the constructor for internal styles")
    @ParameterizedTest(name = "Given name {0}, forced order {1} and internal flag {2} should lead to a valid style")
    @CsvSource({
            "test, 0, false",
            "test2, 777, false",
            "test3, -17, true",
    })
    void constructorTest3(String name, int forceOrder, boolean isInternal) {
        Style style = new Style(name, forceOrder, isInternal);
        assertNotNull(style.getBorder());
        assertNotNull(style.getCellXf());
        assertNotNull(style.getFill());
        assertNotNull(style.getFont());
        assertNotNull(style.getNumberFormat());
        assertEquals(name, style.getName());
        assertEquals(isInternal, style.isInternalStyle());
        assertEquals(forceOrder, style.getInternalID());
    }

    @DisplayName("Test of the Append function on a Border object")
    @Test()
    void appendTest() {
        Style style = new Style();
        Border border = new Border();
        assertEquals(border.hashCode(), style.getBorder().hashCode());
        Border modified = new Border();
        modified.setBottomColor("FFAABBCC");
        modified.setBottomStyle(Border.StyleValue.dashDotDot);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getBorder().hashCode());
    }


    @DisplayName("Test of the Append function on a Font object")
    @Test()
    void appendTest2() {
        Style style = new Style();
        Font font = new Font();
        assertEquals(font.hashCode(), style.getFont().hashCode());
        Font modified = new Font();
        modified.setBold(true);
        modified.setFamily("Arial");
        style.append(modified);
        assertEquals(modified.hashCode(), style.getFont().hashCode());
    }


    @DisplayName("Test of the Append function on a Fill object")
    @Test()
    void appendTest3() {
        Style style = new Style();
        Fill fill = new Fill();
        assertEquals(fill.hashCode(), style.getFill().hashCode());
        Fill modified = new Fill();
        modified.setBackgroundColor("FFAABBCC");
        modified.setForegroundColor("FF112233");
        style.append(modified);
        assertEquals(modified.hashCode(), style.getFill().hashCode());
    }


    @DisplayName("Test of the Append function on a CellXf object")
    @Test()
    void appendTest4() {
        Style style = new Style();
        CellXf cellXf = new CellXf();
        assertEquals(cellXf.hashCode(), style.getCellXf().hashCode());
        CellXf modified = new CellXf();
        modified.setHorizontalAlign(CellXf.HorizontalAlignValue.distributed);
        modified.setTextRotation(35);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCellXf().hashCode());
    }


    @DisplayName("Test of the Append function on a NumberFormat object")
    @Test()
    void appendTest5() {
        Style style = new Style();
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(numberFormat.hashCode(), style.getNumberFormat().hashCode());
        NumberFormat modified = new NumberFormat();
        modified.setNumber(NumberFormat.FormatNumber.format_11);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getNumberFormat().hashCode());
    }


    @DisplayName("Test of the Append function on a combination of all components")
    @Test()
    void appendTest6() {
        Style style = new Style();
        style.getFont().setSize(18f);
        style.getCellXf().setAlignment(CellXf.TextBreakValue.shrinkToFit);
        style.getBorder().setBottomColor("FFAA3344");
        style.getFill().setBackgroundColor("FF55AACC");
        style.getNumberFormat().setCustomFormatID(190);
        Font font = new Font();
        font.setName("Arial");
        CellXf cellXf = new CellXf();
        cellXf.setHorizontalAlign(CellXf.HorizontalAlignValue.justify);
        Border border = new Border();
        border.setTopColor("FF55BB11");
        Fill fill = new Fill();
        fill.setForegroundColor("FFDDDDDD");
        NumberFormat numberFormat = new NumberFormat();
        numberFormat.setCustomFormatCode("##--##");

        style.append(font);
        style.append(cellXf);
        style.append(border);
        style.append(fill);
        style.append(numberFormat);
        assertEquals(18f, style.getFont().getSize());
        assertEquals("Arial", style.getFont().getName());
        assertEquals(CellXf.TextBreakValue.shrinkToFit, style.getCellXf().getAlignment());
        assertEquals(CellXf.HorizontalAlignValue.justify, style.getCellXf().getHorizontalAlign());
        assertEquals("FFAA3344", style.getBorder().getBottomColor());
        assertEquals("FF55BB11", style.getBorder().getTopColor());
        assertEquals("FF55AACC", style.getFill().getBackgroundColor());
        assertEquals("FFDDDDDD", style.getFill().getForegroundColor());
        assertEquals(190, style.getNumberFormat().getCustomFormatID());
        assertEquals("##--##", style.getNumberFormat().getCustomFormatCode());
    }


    @DisplayName("Test of the Append function on a full other style object")
    @Test()
    void appendTest7() {
        Style style = new Style();
        style.getFont().setSize(18f);
        style.getCellXf().setAlignment(CellXf.TextBreakValue.shrinkToFit);
        style.getBorder().setBottomColor("FFAA3344");
        style.getFill().setBackgroundColor("FF55AACC");
        style.getNumberFormat().setCustomFormatID(190);

        Style style2 = new Style();
        style2.getFont().setName("Arial");
        style2.getCellXf().setHorizontalAlign(CellXf.HorizontalAlignValue.justify);
        style2.getBorder().setTopColor("FF55BB11");
        style2.getFill().setForegroundColor("FFDDDDDD");
        style2.getNumberFormat().setCustomFormatCode("##--##");

        style.append(style2);
        assertEquals(18f, style.getFont().getSize());
        assertEquals("Arial", style.getFont().getName());
        assertEquals(CellXf.TextBreakValue.shrinkToFit, style.getCellXf().getAlignment());
        assertEquals(CellXf.HorizontalAlignValue.justify, style.getCellXf().getHorizontalAlign());
        assertEquals("FFAA3344", style.getBorder().getBottomColor());
        assertEquals("FF55BB11", style.getBorder().getTopColor());
        assertEquals("FF55AACC", style.getFill().getBackgroundColor());
        assertEquals("FFDDDDDD", style.getFill().getForegroundColor());
        assertEquals(190, style.getNumberFormat().getCustomFormatID());
        assertEquals("##--##", style.getNumberFormat().getCustomFormatCode());
    }


    @DisplayName("Test of the Append function on a null style component")
    @Test()
    void appendTest8() {
        Style style = new Style();
        style.getBorder().setBottomColor("FFAA6677");
        int hashCode = style.hashCode();
        style.append(null);
        assertEquals(hashCode, style.hashCode());
    }


    @DisplayName("Test of the failing Append function on a invalid style component (null instance)")
    @Test()
    void appendFailTest() {
        Style style = new Style();
        final Style style2a = new Style();
        style.setBorder(null);
        assertThrows(StyleException.class, () -> style2a.append(style));
        final Style style2b = new Style();
        style.setCellXf(null);
        assertThrows(StyleException.class, () -> style2b.append(style));
        final Style style2c = new Style();
        style.setFill(null);
        assertThrows(StyleException.class, () -> style2c.append(style));
        final Style style2d = new Style();
        style.setFont(null);
        assertThrows(StyleException.class, () -> style2d.append(style));
        final Style style2e = new Style();
        style.setNumberFormat(null);
        assertThrows(StyleException.class, () -> style2e.append(style));
    }


    @DisplayName("Test of the failing hashCode function on a invalid style component (null instance)")
    @Test()
    void hashCodeFailTest() {
        final Style styleA = new Style();
        styleA.setBorder(null);
        assertThrows(StyleException.class, () -> styleA.hashCode());
        final Style styleB = new Style();
        styleB.setCellXf(null);
        assertThrows(StyleException.class, () -> styleB.hashCode());
        final Style styleC = new Style();
        styleC.setFill(null);
        assertThrows(StyleException.class, () -> styleC.hashCode());
        final Style styleD = new Style();
        styleD.setFont(null);
        assertThrows(StyleException.class, () -> styleD.hashCode());
        final Style styleE = new Style();
        styleE.setNumberFormat(null);
        assertThrows(StyleException.class, () -> styleE.hashCode());
    }


    @DisplayName("Test of the failing Copy function on a invalid style component (null instance)")
    @Test()
    void copyFailTest() {
        final Style styleA = new Style();
        styleA.setBorder(null);
        assertThrows(StyleException.class, () -> styleA.copy());
        final Style styleB = new Style();
        styleB.setCellXf(null);
        assertThrows(StyleException.class, () -> styleB.copy());
        final Style styleC = new Style();
        styleC.setFill(null);
        assertThrows(StyleException.class, () -> styleC.copy());
        final Style styleD = new Style();
        styleD.setFont(null);
        assertThrows(StyleException.class, () -> styleD.copy());
        final Style styleE = new Style();
        styleE.setNumberFormat(null);
        assertThrows(StyleException.class, () -> styleE.copy());
    }

    // For code coverage
    @DisplayName("Test of the toString function")
    @Test()
    void toStringTest() {
        Style style = new Style();
        String s1 = style.toString();
        style.setName("Test1");
        String s2 = style.toString();
        style.setName(null);
        String s3 = style.toString();
        String hashCode = Integer.toString(style.hashCode());
        assertNotEquals(s1, s2);
        assertTrue(s2.contains("Test1"));
        assertEquals(hashCode, s3);
    }

}
