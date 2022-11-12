package ch.rabanti.nanoxlsx4j.workbooks;

import ch.rabanti.nanoxlsx4j.Workbook;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.styles.AbstractStyle;
import ch.rabanti.nanoxlsx4j.styles.BasicStyles;
import ch.rabanti.nanoxlsx4j.styles.Border;
import ch.rabanti.nanoxlsx4j.styles.CellXf;
import ch.rabanti.nanoxlsx4j.styles.Fill;
import ch.rabanti.nanoxlsx4j.styles.Font;
import ch.rabanti.nanoxlsx4j.styles.NumberFormat;
import ch.rabanti.nanoxlsx4j.styles.Style;
import ch.rabanti.nanoxlsx4j.styles.StyleRepository;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Note: All tests of this class are just for code coverage. The tested
 * functions will be removed in the future
 */
public class ObsoleteTest {
    @DisplayName("Test of the addStyle function (only for code coverage)")
    @Test()
    void addStyleTest() {
        Workbook workbook = new Workbook();
        workbook.addStyle(BasicStyles.Bold());
        assertTrue(StyleRepository.getInstance().getStyles().containsKey(BasicStyles.Bold().hashCode()));
    }

    @DisplayName("Test of the AddStyleComponent function (only for code coverage)")
    @ParameterizedTest(name = "Given... should lead to ")
    @CsvSource(
            {
                    "Border",
                    "CellXf",
                    "Fill",
                    "Font",
                    "NumberFormat",}
    )
    void addStyleComponentTest(String type) {
        Workbook workbook = new Workbook();
        AbstractStyle style = null;
        switch (type) {
            case "Border":
                style = new Border();
                break;
            case "CellXf":
                style = new CellXf();
                break;
            case "Fill":
                style = new Fill();
                break;
            case "Font":
                style = new Font();
                break;
            case "NumberFormat":
                style = new NumberFormat();
                break;
        }
        Style baseStyle = BasicStyles.DottedFill_0_125();
        workbook.addStyleComponent(baseStyle, style);
        assertTrue(StyleRepository.getInstance().getStyles().containsKey(BasicStyles.DottedFill_0_125().hashCode()));
    }

    @DisplayName("Test of the RemoveStyle function with an object (only for code coverage)")
    @Test()
    void removeStyleTest() {
        Workbook workbook = new Workbook();
        Style style = BasicStyles.Bold();
        workbook.addStyle(style);
        workbook.removeStyle(style);
        workbook.removeStyle(style, true);
        workbook.removeStyle(style, false);
        assertTrue(StyleRepository.getInstance().getStyles().containsKey(BasicStyles.Bold().hashCode())); // This is
        // expected
        Style style2 = null;
        assertThrows(StyleException.class, () -> workbook.removeStyle(style2));
    }

    @DisplayName("Test of the RemoveStyle function with a name (only for code coverage)")
    @Test()
    void removeStyleTest2() {
        Workbook workbook = new Workbook();
        Style style = BasicStyles.Bold();
        workbook.addStyle(style);
        workbook.removeStyle(style.getName());
        workbook.removeStyle(style.getName(), true);
        workbook.removeStyle(style.getName(), false);
        assertTrue(StyleRepository.getInstance().getStyles().containsKey(BasicStyles.Bold().hashCode())); // This is
        // expected
        String styleName = null;
        assertThrows(StyleException.class, () -> workbook.removeStyle(styleName));
    }

}
