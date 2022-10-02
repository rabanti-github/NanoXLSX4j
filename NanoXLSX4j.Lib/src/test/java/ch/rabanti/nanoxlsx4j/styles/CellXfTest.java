package ch.rabanti.nanoxlsx4j.styles;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import ch.rabanti.nanoxlsx4j.TestUtils;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;

public class CellXfTest {

	private CellXf exampleStyle;

	public CellXfTest() {
		exampleStyle = new CellXf();
		exampleStyle.setHidden(true);
		exampleStyle.setLocked(true);
		exampleStyle.setForceApplyAlignment(true);
		exampleStyle.setHorizontalAlign(CellXf.HorizontalAlignValue.left);
		exampleStyle.setVerticalAlign(CellXf.VerticalAlignValue.center);
		exampleStyle.setTextDirection(CellXf.TextDirectionValue.horizontal);
		exampleStyle.setAlignment(CellXf.TextBreakValue.shrinkToFit);
		exampleStyle.setTextRotation(75);
		exampleStyle.setIndent(3);
	}

	@DisplayName("Test of the get and set function of the Hidden field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "true", "false", })
	void hiddenTest(boolean value) {
		CellXf cellXf = new CellXf();
		assertFalse(cellXf.isHidden());
		cellXf.setHidden(value);
		assertEquals(value, cellXf.isHidden());
	}

	@DisplayName("Test of the get and set function of the Locked field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "true", "false", })
	void lockedTest(boolean value) {
		CellXf cellXf = new CellXf();
		assertFalse(cellXf.isLocked());
		cellXf.setLocked(value);
		assertEquals(value, cellXf.isLocked());
	}

	@DisplayName("Test of the get and set function of the ForceApplyAlignment field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "true", "false", })
	void forceApplyAlignmentTest(boolean value) {
		CellXf cellXf = new CellXf();
		assertFalse(cellXf.isForceApplyAlignment());
		cellXf.setForceApplyAlignment(value);
		assertEquals(value, cellXf.isForceApplyAlignment());
	}

	@DisplayName("Test of the get and set function of the HorizontalAlign field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "center", "centerContinuous", "distributed", "fill", "general", "justify", "left", "none", "right", })
	void horizontalAlignTest(CellXf.HorizontalAlignValue value) {
		CellXf cellXf = new CellXf();
		assertEquals(CellXf.DEFAULT_HORIZONTAL_ALIGNMENT, cellXf.getHorizontalAlign()); // none is default
		cellXf.setHorizontalAlign(value);
		assertEquals(value, cellXf.getHorizontalAlign());
	}

	@DisplayName("Test of the get and set function of the VerticalAlign field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "bottom", "center", "distributed", "justify", "none", "top", })
	void verticalAlignTest(CellXf.VerticalAlignValue value) {
		CellXf cellXf = new CellXf();
		assertEquals(CellXf.DEFAULT_VERTICAL_ALIGNMENT, cellXf.getVerticalAlign()); // none is default
		cellXf.setVerticalAlign(value);
		assertEquals(value, cellXf.getVerticalAlign());
	}

	@DisplayName("Test of the get and set function of the HorizontalAlign field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "horizontal", "vertical", })
	void textDirectionTest(CellXf.TextDirectionValue value) {
		CellXf cellXf = new CellXf();
		assertEquals(CellXf.DEFAULT_TEXT_DIRECTION, cellXf.getTextDirection()); // horizontal is default
		cellXf.setTextDirection(value);
		assertEquals(value, cellXf.getTextDirection());
		if (value == CellXf.TextDirectionValue.vertical) {
			assertEquals(255, cellXf.getTextRotation());
		}
	}

	@DisplayName("Test of the get and set function of the TextRotation field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "0", "33", "90", "-33", "-90", })
	void textRotationTest(int value) {
		CellXf cellXf = new CellXf();
		assertEquals(0, cellXf.getTextRotation()); // 0 is default
		cellXf.setTextRotation(value);
		assertEquals(value, cellXf.getTextRotation());
	}

	@DisplayName("Test of the failing get and set function of the textRotation field on out-of-range values")
	@ParameterizedTest(name = "GGiven value {0} should lead to an exception")
	@CsvSource({ "91", "-91", "-360", "360", "720", })
	void textRotationFailTest(int value) {
		CellXf cellXf = new CellXf();
		assertEquals(0, cellXf.getTextRotation()); // 0 is default
		assertThrows(FormatException.class, () -> cellXf.setTextRotation(value));
	}

	@DisplayName("Test of the get and set function of the Align field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "none", "shrinkToFit", "wrapText", })
	void alignTest(CellXf.TextBreakValue value) {
		CellXf cellXf = new CellXf();
		assertEquals(CellXf.DEFAULT_ALIGNMENT, cellXf.getAlignment()); // none is default
		cellXf.setAlignment(value);
		assertEquals(value, cellXf.getAlignment());
	}

	@DisplayName("Test of the get and set function of the Indent field")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "0", "1", "99", })
	void indentTest(int value) {
		CellXf cellXf = new CellXf();
		assertEquals(0, cellXf.getIndent()); // 0 is default
		cellXf.setIndent(value);
		assertEquals(value, cellXf.getIndent());
	}

	@DisplayName("Test of the failing set function of the indent field when an invalid value was passed")
	@ParameterizedTest(name = "Given value {0} should lead to the defined field")
	@CsvSource({ "-1", "-999", })
	void indentFailTest(int value) {
		assertThrows(Exception.class, () -> exampleStyle.setIndent(value));
	}

	@DisplayName("Test of the equals method")
	@Test()
	void equalsTest() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		assertTrue(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of Locked)")
	@Test()
	void equalsTest2() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setLocked(false);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of Hidden)")
	@Test()
	void equalsTest2b() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setHidden(false);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of HorizontalAlign)")
	@Test()
	void equalsTest2c() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setHorizontalAlign(CellXf.HorizontalAlignValue.right);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of VerticalAlign)")
	@Test()
	void equalsTest2d() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setVerticalAlign(CellXf.VerticalAlignValue.top);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of ForceApplyAlignment)")
	@Test()
	void equalsTest2e() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setForceApplyAlignment(false);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of TextDirection)")
	@Test()
	void equalsTest2f() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setTextDirection(CellXf.TextDirectionValue.vertical);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of TextRotation)")
	@Test()
	void equalsTest2g() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setTextRotation(27);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of Alignment)")
	@Test()
	void equalsTest2h() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setAlignment(CellXf.TextBreakValue.none);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality of Indent)")
	@Test()
	void equalsTest2i() {
		CellXf style2 = (CellXf) exampleStyle.copy();
		style2.setIndent(77);
		assertFalse(exampleStyle.equals(style2));
	}

	@DisplayName("Test of the equals method (inequality on null or different objects)")
	@ParameterizedTest(name = "Given value should lead to an equal result")
	@CsvSource({ "NULL, ''", "STRING, 'text'", "BOOLEAN, 'true'", })
	void equalsTest3(String sourceType, String sourceValue) {
		Object obj = TestUtils.createInstance(sourceType, sourceValue);
		assertFalse(exampleStyle.equals(obj));
	}

	@DisplayName("Test of the equals method when the origin object is null or not of the same type")
	@ParameterizedTest(name = "Given value {1} should lead to a non-equal result")
	@CsvSource({ "NULL ,''", "BOOLEAN ,'true'", "STRING ,'origin'", })
	void equalsTest5(String sourceType, String sourceValue) {
		Object origin = TestUtils.createInstance(sourceType, sourceValue);
		assertFalse(exampleStyle.equals(origin));
	}

	@DisplayName("Test of the hashCode method (equality of two identical objects)")
	@Test()
	void hashCodeTest() {
		CellXf copy = (CellXf) exampleStyle.copy();
		copy.setInternalID(99); // Should not influence
		assertEquals(exampleStyle.hashCode(), copy.hashCode());
	}

	@DisplayName("Test of the hashCode method (inequality of two different objects)")
	@Test()
	void hashCodeTest2() {
		CellXf copy = (CellXf) exampleStyle.copy();
		copy.setHidden(false);
		assertNotEquals(exampleStyle.hashCode(), copy.hashCode());
	}

	@DisplayName("Test of the compareTo method")
	@Test()
	void compareToTest() {
		CellXf cellXf = new CellXf();
		CellXf other = new CellXf();
		cellXf.setInternalID(null);
		other.setInternalID(null);
		assertEquals(-1, cellXf.compareTo(other));
		cellXf.setInternalID(5);
		assertEquals(1, cellXf.compareTo(other));
		assertEquals(1, cellXf.compareTo(null));
		other.setInternalID(5);
		assertEquals(0, cellXf.compareTo(other));
		other.setInternalID(4);
		assertEquals(1, cellXf.compareTo(other));
		other.setInternalID(6);
		assertEquals(-1, cellXf.compareTo(other));
	}

	// For code coverage

	@DisplayName("Test of the toString function")
	@Test()
	void toStringTest() {
		CellXf cellXf = new CellXf();
		String s1 = cellXf.toString();
		cellXf.setTextRotation(12);
		assertNotEquals(s1, cellXf.toString()); // An explicit value comparison is probably not sensible
	}

	@DisplayName("Test of the getValue function of the HorizontalAlignValue enum")
	@Test()
	void horizontalAlignValueTest() {
		assertEquals(0, CellXf.HorizontalAlignValue.none.getValue());
		assertEquals(1, CellXf.HorizontalAlignValue.left.getValue());
		assertEquals(2, CellXf.HorizontalAlignValue.center.getValue());
		assertEquals(3, CellXf.HorizontalAlignValue.right.getValue());
		assertEquals(4, CellXf.HorizontalAlignValue.fill.getValue());
		assertEquals(5, CellXf.HorizontalAlignValue.justify.getValue());
		assertEquals(6, CellXf.HorizontalAlignValue.general.getValue());
		assertEquals(7, CellXf.HorizontalAlignValue.centerContinuous.getValue());
		assertEquals(8, CellXf.HorizontalAlignValue.distributed.getValue());
	}

	@DisplayName("Test of the getValue function of the VerticalAlignValue enum")
	@Test()
	void verticalAlignValueTest() {
		assertEquals(0, CellXf.VerticalAlignValue.none.getValue());
		assertEquals(1, CellXf.VerticalAlignValue.bottom.getValue());
		assertEquals(2, CellXf.VerticalAlignValue.top.getValue());
		assertEquals(3, CellXf.VerticalAlignValue.center.getValue());
		assertEquals(4, CellXf.VerticalAlignValue.justify.getValue());
		assertEquals(5, CellXf.VerticalAlignValue.distributed.getValue());
	}

	@DisplayName("Test of the getValue function of the TextBreakValue enum")
	@Test()
	void TextBreakValueTest() {
		assertEquals(0, CellXf.TextBreakValue.none.getValue());
		assertEquals(1, CellXf.TextBreakValue.wrapText.getValue());
		assertEquals(2, CellXf.TextBreakValue.shrinkToFit.getValue());
	}

	@DisplayName("Test of the getValue function of the TextDirectionValue enum")
	@Test()
	void TextDirectionValueTest() {
		assertEquals(0, CellXf.TextDirectionValue.horizontal.getValue());
		assertEquals(1, CellXf.TextDirectionValue.vertical.getValue());
	}

}
