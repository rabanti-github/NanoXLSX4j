/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Optional;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

class StyleTest {

    @Test
    @DisplayName("Test of the get and set function of the CurrentBorder property")
    void currentBorderTest() {
        Style style = new Style();
        Border border = new Border();
        assertNotNull(style.getCurrentBorder());
        assertEquals(border.hashCode(), style.getCurrentBorder().hashCode());
        style.setCurrentBorder(border);
        border.setBottomColor("FFAABBCC");
        assertEquals("FFAABBCC", style.getCurrentBorder().getBottomColor());
    }

    @Test
    @DisplayName("Test of the get and set function of the CurrentCellXf property")
    void currentCellXfTest() {
        Style style = new Style();
        CellXf cellXf = new CellXf();
        assertNotNull(style.getCurrentCellXf());
        assertEquals(cellXf.hashCode(), style.getCurrentCellXf().hashCode());
        style.setCurrentCellXf(cellXf);
        cellXf.setIndent(5);
        assertEquals(5, style.getCurrentCellXf().getIndent());
    }

    @Test
    @DisplayName("Test of the get and set function of the CurrentFill property")
    void currentFillTest() {
        Style style = new Style();
        Fill fill = new Fill();
        assertNotNull(style.getCurrentFill());
        assertEquals(fill.hashCode(), style.getCurrentFill().hashCode());
        style.setCurrentFill(fill);
        fill.setBackgroundColor("AACCBBDD");
        assertEquals("AACCBBDD", style.getCurrentFill().getBackgroundColor().getArgbValue());
    }

    @Test
    @DisplayName("Test of the get and set function of the CurrentFont property")
    void currentFontTest() {
        Style style = new Style();
        Font font = new Font();
        assertNotNull(style.getCurrentFont());
        assertEquals(font.hashCode(), style.getCurrentFont().hashCode());
        style.setCurrentFont(font);
        font.setName("Sans Serif");
        assertEquals("Sans Serif", style.getCurrentFont().getName());
    }

    @Test
    @DisplayName("Test of the get and set function of the CurrentNumberFormat property")
    void currentNumberFormatTest() {
        Style style = new Style();
        NumberFormat numberFormat = new NumberFormat();
        assertNotNull(style.getCurrentFill());
        assertEquals(numberFormat.hashCode(), style.getCurrentNumberFormat().hashCode());
        style.setCurrentNumberFormat(numberFormat);
        numberFormat.setNumber(NumberFormat.FormatNumber.FORMAT_15);
        assertEquals(NumberFormat.FormatNumber.FORMAT_15, style.getCurrentNumberFormat().getNumber());
    }

    @Test
    @DisplayName("Test of the get and set function of the Name property")
    void nameTest() {
        Style style = new Style();
        assertEquals(Integer.toString(style.hashCode()), style.getName());
        style.setName("Test");
        assertEquals("Test", style.getName());
    }

    @Test
    @DisplayName("Test of the get function of the IsInternalStyle property")
    void isInternalStyleTest() {
        Style style = new Style();
        assertFalse(style.isInternalStyle());
        Style internalStyle = new Style("test", 0, true);
        assertTrue(internalStyle.isInternalStyle());
    }

    @Test
    @DisplayName("Test of the get and set function of the InternalID property")
    void internalIDTest() {
        Style style = new Style();
        assertTrue(style.getInternalId().isEmpty());
        style.setInternalId(Optional.of(962));
        assertEquals(Optional.of(962), style.getInternalId());
    }

    @Test
    @DisplayName("Test of the default constructor")
    void constructorTest() {
        Style style = new Style();
        assertNotNull(style.getCurrentBorder());
        assertNotNull(style.getCurrentCellXf());
        assertNotNull(style.getCurrentFill());
        assertNotNull(style.getCurrentFont());
        assertNotNull(style.getCurrentNumberFormat());
        assertNotNull(style.getName());
        assertTrue(style.getInternalId().isEmpty());
    }

    @Test
    @DisplayName("Test of the constructor with a name")
    void constructorTest2() {
        Style style = new Style("test1");
        assertNotNull(style.getCurrentBorder());
        assertNotNull(style.getCurrentCellXf());
        assertNotNull(style.getCurrentFill());
        assertNotNull(style.getCurrentFont());
        assertNotNull(style.getCurrentNumberFormat());
        assertEquals("test1", style.getName());
        assertTrue(style.getInternalId().isEmpty());
    }

    @ParameterizedTest
    @DisplayName("Test of the constructor for internal styles")
    @CsvSource({
            "test, 0, false",
            "test2, 777, false",
            "test3, -17, true"
    })
    void constructorTest3(String name, int forceOrder, boolean isInternal) {
        Style style = new Style(name, forceOrder, isInternal);
        assertNotNull(style.getCurrentBorder());
        assertNotNull(style.getCurrentCellXf());
        assertNotNull(style.getCurrentFill());
        assertNotNull(style.getCurrentFont());
        assertNotNull(style.getCurrentNumberFormat());
        assertEquals(name, style.getName());
        assertEquals(isInternal, style.isInternalStyle());
        assertEquals(Optional.of(forceOrder), style.getInternalId());
    }

    @Test
    @DisplayName("Test of the Append function on a Border object")
    void appendTest() {
        Style style = new Style();
        Border border = new Border();
        assertEquals(border.hashCode(), style.getCurrentBorder().hashCode());
        Border modified = new Border();
        modified.setBottomColor("FFAABBCC");
        modified.setBottomStyle(Border.StyleValue.DASH_DOT_DOT);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCurrentBorder().hashCode());
    }

    @Test
    @DisplayName("Test of the Append function on a Font object")
    void appendTest2() {
        Style style = new Style();
        Font font = new Font();
        assertEquals(font.hashCode(), style.getCurrentFont().hashCode());
        Font modified = new Font();
        modified.setBold(true);
        modified.setFamily(Font.FontFamilyValue.MODERN);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCurrentFont().hashCode());
    }

    @Test
    @DisplayName("Test of the Append function on a Fill object")
    void appendTest3() {
        Style style = new Style();
        Fill fill = new Fill();
        assertEquals(fill.hashCode(), style.getCurrentFill().hashCode());
        Fill modified = new Fill();
        modified.setBackgroundColor("FFAABBCC");
        modified.setForegroundColor("FF112233");
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCurrentFill().hashCode());
    }

    @Test
    @DisplayName("Test of the Append function on a CellXf object")
    void appendTest4() {
        Style style = new Style();
        CellXf cellXf = new CellXf();
        assertEquals(cellXf.hashCode(), style.getCurrentCellXf().hashCode());
        CellXf modified = new CellXf();
        modified.setHorizontalAlign(CellXf.HorizontalAlignValue.DISTRIBUTED);
        modified.setTextRotation(35);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCurrentCellXf().hashCode());
    }

    @Test
    @DisplayName("Test of the Append function on a NumberFormat object")
    void appendTest5() {
        Style style = new Style();
        NumberFormat numberFormat = new NumberFormat();
        assertEquals(numberFormat.hashCode(), style.getCurrentNumberFormat().hashCode());
        NumberFormat modified = new NumberFormat();
        modified.setNumber(NumberFormat.FormatNumber.FORMAT_11);
        style.append(modified);
        assertEquals(modified.hashCode(), style.getCurrentNumberFormat().hashCode());
    }

    @Test
    @DisplayName("Test of the Append function on a combination of all components")
    void appendTest6() {
        Style style = new Style();
        style.getCurrentFont().setSize(18f);
        style.getCurrentCellXf().setAlignment(CellXf.TextBreakValue.SHRINK_TO_FIT);
        style.getCurrentBorder().setBottomColor("FFAA3344");
        style.getCurrentFill().setBackgroundColor("FF55AACC");
        style.getCurrentNumberFormat().setCustomFormatId(190);
        Font font = new Font();
        font.setName("Arial");
        CellXf cellXf = new CellXf();
        cellXf.setHorizontalAlign(CellXf.HorizontalAlignValue.JUSTIFY);
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
        assertEquals(18f, style.getCurrentFont().getSize());
        assertEquals("Arial", style.getCurrentFont().getName());
        assertEquals(CellXf.TextBreakValue.SHRINK_TO_FIT, style.getCurrentCellXf().getAlignment());
        assertEquals(CellXf.HorizontalAlignValue.JUSTIFY, style.getCurrentCellXf().getHorizontalAlign());
        assertEquals("FFAA3344", style.getCurrentBorder().getBottomColor());
        assertEquals("FF55BB11", style.getCurrentBorder().getTopColor());
        assertEquals("FF55AACC", style.getCurrentFill().getBackgroundColor().getArgbValue());
        assertEquals("FFDDDDDD", style.getCurrentFill().getForegroundColor().getArgbValue());
        assertEquals(190, style.getCurrentNumberFormat().getCustomFormatId());
        assertEquals("##--##", style.getCurrentNumberFormat().getCustomFormatCode());
    }

    @Test
    @DisplayName("Test of the Append function on a full other style object")
    void appendTest7() {
        Style style = new Style();
        style.getCurrentFont().setSize(18f);
        style.getCurrentCellXf().setAlignment(CellXf.TextBreakValue.SHRINK_TO_FIT);
        style.getCurrentBorder().setBottomColor("FFAA3344");
        style.getCurrentFill().setBackgroundColor("FF55AACC");
        style.getCurrentNumberFormat().setCustomFormatId(190);

        Style style2 = new Style();
        style2.getCurrentFont().setName("Arial");
        style2.getCurrentCellXf().setHorizontalAlign(CellXf.HorizontalAlignValue.JUSTIFY);
        style2.getCurrentBorder().setTopColor("FF55BB11");
        style2.getCurrentFill().setForegroundColor("FFDDDDDD");
        style2.getCurrentNumberFormat().setCustomFormatCode("##--##");

        style.append(style2);
        assertEquals(18f, style.getCurrentFont().getSize());
        assertEquals("Arial", style.getCurrentFont().getName());
        assertEquals(CellXf.TextBreakValue.SHRINK_TO_FIT, style.getCurrentCellXf().getAlignment());
        assertEquals(CellXf.HorizontalAlignValue.JUSTIFY, style.getCurrentCellXf().getHorizontalAlign());
        assertEquals("FFAA3344", style.getCurrentBorder().getBottomColor());
        assertEquals("FF55BB11", style.getCurrentBorder().getTopColor());
        assertEquals("FF55AACC", style.getCurrentFill().getBackgroundColor().getArgbValue());
        assertEquals("FFDDDDDD", style.getCurrentFill().getForegroundColor().getArgbValue());
        assertEquals(190, style.getCurrentNumberFormat().getCustomFormatId());
        assertEquals("##--##", style.getCurrentNumberFormat().getCustomFormatCode());
    }

    @Test
    @DisplayName("Test of the Append function on a null style component")
    void appendTest8() {
        Style style = new Style();
        style.getCurrentBorder().setBottomColor("FFAA6677");
        int hashCode = style.hashCode();
        style.append(null);
        assertEquals(hashCode, style.hashCode());
    }

    @Test
    @DisplayName("Test of the failing Append function on a invalid style component (null instance)")
    void appendFailTest() {
        Style style = new Style();
        Style style2a = new Style();
        style.setCurrentBorder(null);
        assertThrows(StyleException.class, () -> style2a.append(style));
        Style style2b = new Style();
        style.setCurrentCellXf(null);
        assertThrows(StyleException.class, () -> style2b.append(style));
        Style style2c = new Style();
        style.setCurrentFill(null);
        assertThrows(StyleException.class, () -> style2c.append(style));
        Style style2d = new Style();
        style.setCurrentFont(null);
        assertThrows(StyleException.class, () -> style2d.append(style));
        Style style2e = new Style();
        style.setCurrentNumberFormat(null);
        assertThrows(StyleException.class, () -> style2e.append(style));
    }

    @Test
    @DisplayName("Test of the failing GetHashCode function on a invalid style component (null instance)")
    void getHashCodeFailTest() {
        Style styleA = new Style();
        styleA.setCurrentBorder(null);
        assertThrows(StyleException.class, styleA::hashCode);
        Style styleB = new Style();
        styleB.setCurrentCellXf(null);
        assertThrows(StyleException.class, styleB::hashCode);
        Style styleC = new Style();
        styleC.setCurrentFill(null);
        assertThrows(StyleException.class, styleC::hashCode);
        Style styleD = new Style();
        styleD.setCurrentFont(null);
        assertThrows(StyleException.class, styleD::hashCode);
        Style styleE = new Style();
        styleE.setCurrentNumberFormat(null);
        assertThrows(StyleException.class, styleE::hashCode);
    }

    @Test
    @DisplayName("Test of the failing Copy function on a invalid style component (null instance)")
    void copyFailTest() {
        Style styleA = new Style();
        styleA.setCurrentBorder(null);
        assertThrows(StyleException.class, styleA::copy);
        Style styleB = new Style();
        styleB.setCurrentCellXf(null);
        assertThrows(StyleException.class, styleB::copy);
        Style styleC = new Style();
        styleC.setCurrentFill(null);
        assertThrows(StyleException.class, styleC::copy);
        Style styleD = new Style();
        styleD.setCurrentFont(null);
        assertThrows(StyleException.class, styleD::copy);
        Style styleE = new Style();
        styleE.setCurrentNumberFormat(null);
        assertThrows(StyleException.class, styleE::copy);
    }

    // For code coverage
    @Test
    @DisplayName("Test of the ToString function")
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
        assertTrue(s3.contains(hashCode));
    }
}
