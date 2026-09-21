package ch.rabanti.nanoxlsx4j;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.util.Date;
import java.util.Optional;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.NullSource;
import org.junit.jupiter.params.provider.ValueSource;
import ch.rabanti.nanoxlsx4j.enums.FormulaError;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;

public class DefinedNameTest {

    @ParameterizedTest
    @DisplayName("Test of valid defined name identifiers")
    @ValueSource(
            strings = {
                    "Revenue_2026",
                    "_private.name",
                    "\\legacy",
                    "Übersicht",
                    "工作表Name"}
    )
    void validNameTest(String name) {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName definedName = new DefinedName(workbook, DefinedName.NameType.CONSTANT, name, 1, null);

        assertEquals(name, definedName.getName());
    }

    @ParameterizedTest
    @DisplayName("Test of invalid defined name identifiers")
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    "   ",
                    "1Name",
                    ".Name",
                    "Bad Name",
                    "Bad-Name",
                    "C",
                    "r",
                    "A1",
                    "XFD1048576"}
    )
    void invalidNameTest(String name) {
        Workbook workbook = new Workbook("Sheet1");

        assertThrows(
                FormatException.class,
                () -> workbook.addDefinedNameConstant(name, 1)
        );
    }

    @Test
    @DisplayName("Test of the maximum defined name length")
    void nameLengthTest() {
        Workbook workbook = new Workbook("Sheet1");
        assertNotNull(new DefinedName(workbook, DefinedName.NameType.CONSTANT, "N".repeat(255), 1, null));
        assertThrows(
                FormatException.class,
                () -> new DefinedName(workbook, DefinedName.NameType.CONSTANT, "N".repeat(256), 1, null)
        );
    }

    @ParameterizedTest
    @DisplayName("Test of AddDefinedNameCell overloads")
    @ValueSource(
            ints = {
                    0, // Address
                    1, // row, column
                    2} // address object
    )
    void addDefinedNameCellTest(int overload) {
        Workbook workbook = new Workbook("Target");
        Worksheet target = workbook.getCurrentWorksheet();
        DefinedName name = switch (overload) {
            case 0 -> workbook.addDefinedNameCell("CellName", target, "b2", null, "comment");
            case 1 -> workbook.addDefinedNameCell("CellName", target, 1, 1, null, "comment");
            default -> workbook.addDefinedNameCell("CellName", target, new Address("B2"), null, "comment");
        };

        assertEquals(DefinedName.NameType.CELL, name.getType());
        assertEquals("$B$2", name.getTextValue());
        assertEquals(new Address("$B$2"), name.getValue());
        assertSame(target, name.getTargetWorksheet());
        assertNull(name.getLocalSheet());
        assertEquals("comment", name.getComment());
        assertEquals(FormulaError.NO_ERROR, name.getError());
        assertFalse(name.hasExternalReferences());
    }

    @ParameterizedTest
    @DisplayName("Test of AddDefinedNameRange overloads")
    @ValueSource(
            ints = {
                    0,
                    1,
                    2,
                    3}
    )
    void addDefinedNameRangeTest(int overload) {
        Workbook workbook = new Workbook("Target");
        Worksheet target = workbook.getCurrentWorksheet();
        DefinedName name = switch (overload) {
            case 0 -> workbook.addDefinedNameRange("RangeName", target, "a1:b3");
            case 1 -> workbook.addDefinedNameRange("RangeName", target, new Address("A1"), new Address("B3"));
            case 2 -> workbook.addDefinedNameRange("RangeName", target, 0, 0, 1, 2);
            default -> workbook.addDefinedNameRange("RangeName", target, new Range("A1:B3"));
        };

        assertEquals(DefinedName.NameType.RANGE, name.getType());
        assertEquals("$A$1:$B$3", name.getTextValue());
        assertEquals(new Range("$A$1:$B$3"), name.getValue());
        assertSame(target, name.getTargetWorksheet());
    }

    @ParameterizedTest
    @DisplayName("Test of all supported defined name constant types")
    @ValueSource(
            strings = {
                    "String",
                    "Whitespace",
                    "Bool",
                    "Byte",
                    "SByte",
                    "Decimal",
                    "Double",
                    "Float",
                    "Int",
                    "UInt",
                    "Long",
                    "ULong",
                    "Short",
                    "UShort",
                    "DateTime",
                    "TimeSpan",
                    "Object"
            }
    )
    void addDefinedNameConstantTest(String kind) {
        Object value = createConstant(kind);
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = new DefinedName(
                workbook, DefinedName.NameType.CONSTANT, "ConstantName", value, null
        );

        assertEquals(DefinedName.NameType.CONSTANT, name.getType());
        assertSame(value, name.getValue());
        assertNotNull(name.getTextValue());
        assertNull(name.getTargetWorksheet());
    }

    @ParameterizedTest
    @DisplayName("Test of invalid defined name values (formula)")
    @NullSource
    @ValueSource(
            strings = {
                    "",
                    " "}
    )
    void invalidFormulaTest(String formula) {
        Workbook workbook = new Workbook("Sheet1");
        assertThrows(WorksheetException.class, () -> workbook.addDefinedNameFormula("FormulaName", formula));
    }

    @Test
    @DisplayName("Test of formula and null constant defined names")
    void formulaAndNullConstantTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName formula = workbook.addDefinedNameFormula("FormulaName", "SUM(1,2)", null, "note");
        assertEquals(DefinedName.NameType.FORMULA, formula.getType());
        assertEquals("SUM(1,2)", formula.getTextValue());
        assertEquals("note", formula.getComment());
        assertFalse(formula.hasExternalReferences());
        assertThrows(WorksheetException.class, () -> workbook.addDefinedNameConstant("NullName", null));
        assertThrows(FormatException.class, () -> workbook.addDefinedNameConstant("EmptyName", ""));
    }

    @ParameterizedTest
    @DisplayName("Test of external workbook reference detection in added defined-name formulas")
    @ValueSource(
            strings = {
                    "[book.xlsx]Sheet1!A1",
                    "'..\\[book.xlsx]Data'!$B$2",
                    "'../[book.xlsx]Data'!$B$2",
                    "'C:\\temp\\[book one.xlsx]Sheet 1'!$A$1",
                    "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)",
                    "[1]Sheet1!$A$1"
            }
    )
    void addDefinedNameFormulaExternalReferenceTest(String expression) {
        Workbook workbook = new Workbook("Sheet1");

        DefinedName name = new DefinedName(
                workbook, DefinedName.NameType.FORMULA, "ExternalFormula", expression, null
        );

        assertTrue(name.hasExternalReferences());
        assertEquals(expression, name.getTextValue());
        assertEquals(expression, name.getValue());
    }

    @ParameterizedTest
    @DisplayName("Test of non-external defined-name formulas")
    @ValueSource(
            strings = {
                    "SUM(A1:A2)",
                    "Table1[Column]",
                    "R[1]C[1]",
                    "INDIRECT(\"[1]Sheet1!A1\")"}
    )
    void addDefinedNameFormulaWithoutExternalReferenceTest(String expression) {
        Workbook workbook = new Workbook("Sheet1");

        DefinedName name = new DefinedName(
                workbook, DefinedName.NameType.FORMULA, "LocalFormula", expression, null
        );

        assertFalse(name.hasExternalReferences());
    }

    @Test
    @DisplayName("Test of invalid cell and range defined names")
    void invalidCellAndRangeTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet worksheet = workbook.getCurrentWorksheet();
        assertThrows(
                WorksheetException.class,
                () -> workbook.addDefinedNameCell("Name1", null, new Address("A1"))
        );
        assertThrows(
                WorksheetException.class,
                () -> workbook.addDefinedNameCell("Name2", worksheet, (Address) null)
        );
        assertThrows(
                WorksheetException.class,
                () -> workbook.addDefinedNameRange("Name3", null, new Range("A1:B2"))
        ); // In contrast to C# this is a Workbook exception
        assertThrows(
                FormatException.class,
                () -> workbook.addDefinedNameRange("Name4", worksheet, (Range) null)
        );
        assertThrows(
                FormatException.class,
                () -> workbook.addDefinedNameCell("Name5", worksheet, "A1:B2")
        );
        assertEquals("$A$1:$A$1", workbook.addDefinedNameRange("Name6", worksheet, "A1").getTextValue());
    }

    @Test
    @DisplayName("Test of case-insensitive identity and defined name scopes")
    void scopeAndIdentityTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet sheet1 = workbook.getCurrentWorksheet();
        workbook.addWorksheet("Sheet2");
        Worksheet sheet2 = workbook.getCurrentWorksheet();
        DefinedName global = workbook.addDefinedNameConstant("Rate", 1);
        DefinedName local1 = workbook.addDefinedNameConstant("Rate", 2, sheet1);
        DefinedName local2 = workbook.addDefinedNameConstant("RATE", 3, sheet2);

        assertSame(global, workbook.getDefinedName("rate", null));
        assertSame(local1, workbook.getDefinedName("RATE", sheet1));
        assertSame(local2, workbook.getDefinedName("rate", sheet2));
        assertNull(workbook.getDefinedName("missing", null));
        assertThrows(WorksheetException.class, () -> workbook.addDefinedNameConstant("rAtE", 4));
        assertThrows(WorksheetException.class, () -> workbook.addDefinedNameConstant("rAtE", 4, sheet1));
        assertEquals(3, workbook.getDefinedNames().size());
    }

    @Test
    @DisplayName("Test of removing a defined name and invalidating formula references")
    void removeDefinedNameTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName removed = workbook.addDefinedNameConstant("RemovedName", 5);
        DefinedName retained = workbook.addDefinedNameConstant("RetainedName", 6);
        workbook.getCurrentWorksheet().addCellReference(removed, "A1");
        workbook.getCurrentWorksheet().addCellReference(retained, "A2");
        workbook.addWorksheet("Sheet2");
        workbook.getCurrentWorksheet().addCellReference(removed, "B1");

        assertTrue(workbook.removeDefinedName("removedname"));
        assertNull(workbook.getWorksheets().get(0).getCell(new Address("A1")).getFormula().getDefinedNameReference());
        assertNull(workbook.getWorksheets().get(1).getCell(new Address("B1")).getFormula().getDefinedNameReference());
        assertSame(
                retained,
                workbook.getWorksheets().get(0).getCell(new Address("A2")).getFormula().getDefinedNameReference()
        );
        assertFalse(workbook.removeDefinedName("missing"));
    }

    @Test
    @DisplayName("Test of removing all defined names and invalidating formula references")
    void removeAllDefinedNameTest() {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name1 = workbook.addDefinedNameConstant("name1", 5);
        DefinedName name2 = workbook.addDefinedNameConstant("name2", 6);
        workbook.getCurrentWorksheet().addCellReference(name1, "A1");
        workbook.getCurrentWorksheet().addCellReference(name2, "A2");
        workbook.addWorksheet("Sheet2");
        workbook.getCurrentWorksheet().addCellReference(name1, "B1");

        workbook.clearDefinedNames();

        assertNull(workbook.getWorksheets().get(0).getCell(new Address("A1")).getFormula().getDefinedNameReference());
        assertNull(workbook.getWorksheets().get(0).getCell(new Address("A2")).getFormula().getDefinedNameReference());
        assertNull(workbook.getWorksheets().get(1).getCell(new Address("B1")).getFormula().getDefinedNameReference());
        assertTrue(workbook.getDefinedNames().isEmpty());
        assertNotNull(workbook.getDefinedNames());
    }

    @Test
    @DisplayName("Test of DefinedName equality, hashing and object equality")
    void equalityTest() {
        Workbook workbook1 = new Workbook("Sheet1");
        Workbook workbook2 = new Workbook("Sheet1");
        DefinedName a = new DefinedName(
                workbook1, DefinedName.NameType.CONSTANT, "CaseName", 1, null, null, "comment", Optional.empty()
        );
        DefinedName b = new DefinedName(
                workbook2, DefinedName.NameType.CONSTANT, "casename", 1, null, null, "comment", Optional.empty()
        );

        assertTrue(a.equals(a));
        assertTrue(a.equals(b));
        assertTrue(a.equals((Object) b));
        assertEquals(a.hashCode(), b.hashCode());
        assertFalse(a.equals((DefinedName) null));
        assertFalse(a.equals((Object) null));
        assertFalse(a.equals("wrong"));
        assertFalse(a.equals(new DefinedName(
                workbook2, DefinedName.NameType.CONSTANT, "OtherName", 1, null
        )));
    }

    @Test
    @DisplayName("Test of DefinedName comparison and string representation")
    void compareToAndToStringTest() {
        Workbook workbook = new Workbook("Sheet1");
        Worksheet sheet = workbook.getWorksheets().get(0);
        DefinedName a = new DefinedName(workbook, DefinedName.NameType.CONSTANT, "Alpha", 1, null);
        DefinedName sameNameLocal = new DefinedName(
                workbook, DefinedName.NameType.CONSTANT, "alpha", 1, null, sheet
        );
        DefinedName z = new DefinedName(workbook, DefinedName.NameType.FORMULA, "Zulu", "1+1", null);

        assertEquals(1, a.compareTo(null));
        assertTrue(a.compareTo(z) < 0);
        assertTrue(a.compareTo(sameNameLocal) < 0);
        assertTrue(sameNameLocal.compareTo(a) > 0);
        assertTrue(a.toString().contains("name=Alpha"));
        assertTrue(a.toString().contains("scope=workbook"));
        assertTrue(sameNameLocal.toString().contains("sheet:Sheet1"));
    }

    @Test
    @DisplayName("Test of every DefinedName comparison component")
    void compareToComponentsTest() {
        Workbook workbook1 = new Workbook("Sheet1");
        Workbook workbook2 = new Workbook("Other");
        Worksheet first = workbook1.getWorksheets().get(0);
        Worksheet second = new Worksheet("Sheet2", 2, workbook1);
        first.setSheetId(1);
        second.setSheetId(2);
        DefinedName baseline = new DefinedName(
                workbook1, DefinedName.NameType.CONSTANT, "Name", 1, null, first, "a", Optional.empty()
        );

        assertNotEquals(
                0, baseline.compareTo(new DefinedName(
                        workbook2, DefinedName.NameType.FORMULA, "Name", "1", null, first, "a", Optional.empty()
                ))
        );
        assertTrue(baseline.compareTo(new DefinedName(
                workbook2, DefinedName.NameType.CONSTANT, "Name", 1, null, second, "a", Optional.empty()
        )) < 0);
        assertEquals(
                0, baseline.compareTo(new DefinedName(
                        workbook2, DefinedName.NameType.CONSTANT, "name", 1, null, first, "a", Optional.empty()
                ))
        );
        assertTrue(baseline.compareTo(new DefinedName(
                workbook2, DefinedName.NameType.CONSTANT, "Name", 2, null, first, "a", Optional.empty()
        )) < 0);
        assertTrue(baseline.compareTo(new DefinedName(
                workbook2, DefinedName.NameType.CONSTANT, "Name", 1, null, first, "b", Optional.empty()
        )) < 0);

        DefinedName target1 = new DefinedName(
                workbook1, DefinedName.NameType.CELL, "Target", "A1", first, null
        );
        DefinedName target2 = new DefinedName(
                workbook2, DefinedName.NameType.CELL, "Target", "A1", second, null
        );
        assertTrue(target1.compareTo(target2) < 0);
    }

    @ParameterizedTest
    @DisplayName("Test of resolving defined names read from workbook XML")
    @CsvSource(
            delimiter = '|',
            quoteCharacter = '"',
            value = {
                    "\"\"\"text\"\"\"|CONSTANT|text",
                    "TRUE|CONSTANT|TRUE",
                    "42|CONSTANT|42",
                    "2.5|CONSTANT|2.5",
                    "'Sheet1'!$A$1|CELL|$A$1",
                    "'Sheet1'!$A$1:$B$2|RANGE|$A$1:$B$2",
                    "SUM(1,2)|FORMULA|SUM(1,2)",
                    "#REF!|FORMULA|#REF!",
                    "'Missing'!$A$1|CELL|$A$1",
                    "'Sheet1'!invalid|FORMULA|'Sheet1'!invalid"
            }
    )
    void resolveDefinedNameTest(
            String reference, DefinedName.NameType expectedType, String expectedText
    ) {
        Workbook workbook = new Workbook("Sheet1");
        DefinedName name = DefinedName.resolveDefinedName("ResolvedName", reference, workbook, null, "comment");

        assertEquals(expectedType, name.getType());
        assertEquals(expectedText, name.getTextValue());
        assertEquals("comment", name.getComment());
        assertEquals(reference.equals("#REF!") ? FormulaError.REFERENCE : FormulaError.NO_ERROR, name.getError());
    }

    @ParameterizedTest
    @DisplayName("Test external references in resolved defined names")
    @ValueSource(
            strings = {
                    "'[1]Sheet1'!$A$1",
                    "SUM([1]Sheet1!$A$1)",
                    "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1)"
            }
    )
    void resolveDefinedNameExternalReferenceTest(String reference) {
        Workbook workbook = new Workbook("Sheet1");

        DefinedName name = DefinedName.resolveDefinedName(
                "ResolvedExternalName", reference, workbook, null, null
        );

        assertTrue(name.hasExternalReferences());
    }

    @Test
    @DisplayName("Test that DefinedName requires a workbook")
    void nullWorkbookTest() {
        assertThrows(
                FormatException.class,
                () -> new DefinedName(null, DefinedName.NameType.CONSTANT, "Name", 1, null)
        );
        assertThrows(
                FormatException.class,
                () -> new DefinedName(new Workbook("Sheet1"), DefinedName.NameType.FORMULA, "Name", "", null)
        );
    }

    @ParameterizedTest
    @DisplayName("Test of the ReplaceExpression method")
    @CsvSource(
            delimiter = '|',
            value = {
                    "$A$1-$A$2|$B$2-$C$2",
                    "A1|B1",
                    "x|x",
                    "A|a",
                    "[1]worksheet1!$C$1|[extWorkbook.xlsx]worksheet1!$C$1"
            }
    )
    void replaceExpressionTest(String oldExpression, String newExpression) {
        Workbook workbook = new Workbook("sheet1");
        DefinedName name = new DefinedName(
                workbook, DefinedName.NameType.FORMULA, "name", oldExpression, null, null, "comment", Optional.empty()
        );
        assertEquals(oldExpression, name.getTextValue());

        name.ReplaceExpression(newExpression);

        assertEquals(newExpression, name.getTextValue());
        assertEquals(DefinedName.NameType.FORMULA, name.getType());
        assertEquals("comment", name.getComment());
    }

    @ParameterizedTest
    @DisplayName("Test of the ignoring ReplaceExpression method on incompatible types")
    @CsvSource(
            delimiter = '|',
            value = {
                    "CONSTANT|A",
                    "CELL|$B$2",
                    "RANGE|$A$1:$C$2"
            }
    )
    void replaceExpressionIgnoreTest(DefinedName.NameType type, String expression) {
        Workbook workbook = new Workbook("sheet1");
        Worksheet targetWorksheet = type == DefinedName.NameType.CONSTANT
                ? null
                : workbook.getWorksheets().get(0);
        DefinedName name = new DefinedName(
                workbook, type, "name", expression, targetWorksheet, null, "comment", Optional.empty()
        );
        assertEquals(expression, name.getTextValue());

        name.ReplaceExpression("newExpression");

        assertEquals(expression, name.getTextValue());
        assertEquals(type, name.getType());
        assertEquals("comment", name.getComment());
    }

    private static Object createConstant(String kind) {
        // Java has no unsigned primitive types; use types that preserve the C# reference values.
        return switch (kind) {
            case "String" -> "A \"quoted\" string";
            case "Whitespace" -> "   ";
            case "Bool" -> true;
            case "Byte" -> (byte) 1;
            case "SByte" -> (byte) -2;
            case "Decimal" -> new BigDecimal("3.25");
            case "Double" -> 4.5d;
            case "Float" -> 5.75f;
            case "Int" -> -6;
            case "UInt" -> 7L;
            case "Long" -> -8L;
            case "ULong" -> BigInteger.valueOf(9);
            case "Short" -> (short) -10;
            case "UShort" -> 11;
            case "DateTime" -> Date.from(LocalDateTime.of(2026, 7, 31, 12, 30).toInstant(ZoneOffset.UTC));
            case "TimeSpan" -> Duration.ofMinutes(810);
            default -> new ConstantObject();
        };
    }

    private static final class ConstantObject {
        @Override
        public String toString() {
            return "custom";
        }
    }
}
