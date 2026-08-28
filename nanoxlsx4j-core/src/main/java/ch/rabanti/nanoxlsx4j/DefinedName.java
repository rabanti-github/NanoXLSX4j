/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j;

import java.util.Comparator;
import java.util.HashSet;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Pattern;

import ch.rabanti.nanoxlsx4j.enums.Errors;
import ch.rabanti.nanoxlsx4j.enums.FormulaError;
import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;
import ch.rabanti.nanoxlsx4j.utils.Validators;

/**
 * Class representing a defined name within a workbook. A defined name is a descriptive text that represents a cell, a
 * range of cells, a formula, or a constant value, and can be referenced from formulas in worksheets (e.g.
 * {@code =SUM(MyRange)}).
 *
 * <p>Remarks: A defined name has a workbook scope by default ({@link DefinedName#getLocalSheet()} is null). When a
 * {@link Worksheet} is supplied as {@link DefinedName#getLocalSheet()}, the defined name is scoped to that worksheet
 * (corresponding to the {@code localSheetId} attribute in the OOXML representation). The
 * {@link DefinedName#getTextValue()} is stored verbatim — NanoXLSX does not parse or evaluate it.
 * </p>
 */
public class DefinedName implements Comparable<DefinedName> {

//enums

    /**
     * Enum to specify the type of the defined name
     */
    public enum NameType {
        /** Defined name is a single cell */
        CELL,
        /** Defined name is a cell range */
        RANGE,
        /** Defined name is a formula */
        FORMULA,
        /** Defined name is a constant value */
        CONSTANT

    }

    //constants
    private static final Pattern EXT_WORKSHEET_REFERENCE_REGEX = Pattern.compile("^\\[[0-9]+\\].+");
    private static final int MAX_NAME_LENGTH = 255;
    /**
     * Disallowed names for defined names (ignore case)
     */
    private static final Set<String> DISALLOWED_NAMES;

    static {
        DISALLOWED_NAMES = new HashSet<String>();
        DISALLOWED_NAMES.add("C");
        DISALLOWED_NAMES.add("R");
    }

    /**
     * Allowed special characters at the start of a defined name
     */
    private static final char[] ALLOWED_NAME_START_CHARS = {
            '\\',
            '_'};
    /**
     * Allowed special characters after the first character of a defined name. \remark <remarks>'\' is accepted because
     * Excel allows it, although it is not documented in the official naming rules.</remarks>
     *
     * <p>Remarks: '\' is accepted because Excel allows it, although it is not documented in the official naming
     * rules.</p>
     */
    private static final char[] ALLOWED_NAME_CHARS = {
            '_',
            '.',
            '\\'};

    //privateFields
    private String name;
    private NameType type;
    private Object value;
    private Worksheet localSheet;
    private Worksheet targetWorksheet;
    private String textValue;
    private String comment;
    private FormulaError error;
    boolean externalReferences;
    private FeatureSet features;

//getters&setters

    /**
     * Type of the defined name
     */
    public NameType getType() {
        return type;
    }

    /**
     * Gets the name of the defined name as it appears in the workbook (e.g. {@code MyRange}).
     */
    public String getName() {
        return name;
    }

    /**
     * Gets the target worksheet in case of Cell or Range values. For other types, like formulas or constants, the
     * target worksheet is null
     */
    public Worksheet getTargetWorksheet() {
        return targetWorksheet;
    }

    /**
     * Gets the textual reference of the defined name. This is stored verbatim and may be a cell address (e.g.
     * {@code $A$1}), a range (e.g. {@code $A$1:$A$10}), a formula (e.g. {@code SUM(Sheet1!$A$1:$A$10)}), or a constant
     * value.
     *
     * <p>Remarks: Do not add the target worksheet name (e.g. {@code Sheet1}) in front of the Reference in case of
     * cells or ranges. The worksheet is automatically added by the defined {@link DefinedName#getTargetWorksheet()}
     * </p>
     */
    public String getTextValue() {
        return textValue;
    }

    /**
     * Gets the raw reference of the defined Name. The value will be transformed in its appropriate text value
     * ({@link DefinedName#getTextValue()}). If the object is not supported (integer, float, date / date and time,
     * boolean or string), toString() will be used to determine the text value
     */
    public Object getValue() {
        return value;
    }

    /**
     * Gets the worksheet that scopes (constraint) this defined name. If null, the defined name has workbook scope and
     * is visible from any worksheet. If non-null, the defined name is scoped to the referenced worksheet (mapped to the
     * {@code localSheetId} attribute on save).
     */

    public Worksheet getLocalSheet() {
        return localSheet;
    }

    /**
     * Gets the optional comment associated with the defined name. Maps to the {@code comment} attribute in OOXML. May
     * be null when no comment was set.
     */
    public String getComment() {
        return comment;
    }

    /**
     * Gets a possible error of the whole value in a defined name. Default is {@link FormulaError#NO_ERROR}.
     *
     * <p>Remarks: Errors within formula expressions are not set to an error in this property. The default will be
     * {@link FormulaError#NO_ERROR}.</p>
     */
    public FormulaError getError() {
        return error;
    }

    /**
     * Gets whether the value contains a reference or multiple references to an external source (e.g. an external
     * workbook)
     *
     * @return If true, the defined name contains external references
     */
    public boolean hasExternalReferences() {
        return externalReferences;
    }

    /**
     * Gets the feature set of the defined name
     *
     * @return Feature set
     */
    public FeatureSet getFeatures() {
        return features;
    }

//constructors

    /**
     * Constructs a new defined name.
     *
     * <p>Remarks: Use {@link Workbook#addDefinedNameCell(String, Worksheet, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameConstant(String, Object, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)} (or available overloaded methods) to
     * conveniently add defined names</p>
     *
     * @param workbook  Workbook reference (to check for duplicate names)
     * @param type      Type of the defined name
     * @param name      Name of the defined name. Must be non-empty, must not start with a digit, and must not match a
     *                  cell reference in the range A1 - XFD1048576.
     * @param reference Reference text (cell, range, formula, or constant). Must be non-empty.
     * @param worksheet Target worksheet in case of cells or cell ranges
     * @throws FormatException    Thrown if {@code workbook} is null.
     * @throws FormatException    Thrown if {@code name} is null, empty, only whitespaces, starts with a digit or
     *                            contains illegal characters.
     * @throws FormatException    Thrown if {@code reference} is null or resolved to an empty string
     * @throws WorksheetException Thrown if {@code reference} already exists or matches a cell reference in the same
     *                            scope
     */
    DefinedName(Workbook workbook, NameType type, String name, Object reference, Worksheet worksheet) {
        this(workbook, type, name, reference, worksheet, null, null, Optional.empty());
    }

    /**
     * Constructs a new defined name.
     *
     * <p>Remarks: Use {@link Workbook#addDefinedNameCell(String, Worksheet, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameConstant(String, Object, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)} (or available overloaded methods) to
     * conveniently add defined names</p>
     *
     * @param workbook   Workbook reference (to check for duplicate names)
     * @param type       Type of the defined name
     * @param name       Name of the defined name. Must be non-empty, must not start with a digit, and must not match a
     *                   cell reference in the range A1 - XFD1048576.
     * @param reference  Reference text (cell, range, formula, or constant). Must be non-empty.
     * @param worksheet  Target worksheet in case of cells or cell ranges
     * @param localSheet Optional worksheet that scopes the defined name. Pass null for workbook scope.
     * @throws FormatException    Thrown if {@code workbook} is null.
     * @throws FormatException    Thrown if {@code name} is null, empty, only whitespaces, starts with a digit or
     *                            contains illegal characters.
     * @throws FormatException    Thrown if {@code reference} is null or resolved to an empty string
     * @throws WorksheetException Thrown if {@code reference} already exists or matches a cell reference in the same
     *                            scope
     */
    DefinedName(
            Workbook workbook, NameType type, String name, Object reference, Worksheet worksheet,
            Worksheet localSheet
    ) {
        this(workbook, type, name, reference, worksheet, localSheet, null, Optional.empty());
    }

    /**
     * Constructs a new defined name.
     *
     * <p>Remarks: Use {@link Workbook#addDefinedNameCell(String, Worksheet, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)},
     * {@link Workbook#addDefinedNameConstant(String, Object, Worksheet, String)},
     * {@link Workbook#addDefinedNameFormula(String, String, Worksheet, String)} (or available overloaded methods) to
     * conveniently add defined names</p>
     *
     * @param workbook              Workbook reference (to check for duplicate names)
     * @param type                  Type of the defined name
     * @param name                  Name of the defined name. Must be non-empty, must not start with a digit, and must
     *                              not match a cell reference in the range A1 - XFD1048576.
     * @param reference             Reference text (cell, range, formula, or constant). Must be non-empty.
     * @param worksheet             Target worksheet in case of cells or cell ranges
     * @param localSheet            Optional worksheet that scopes the defined name. Pass null for workbook scope.
     * @param comment               Optional comment.
     * @param containsExternalLinks Optional pre-resolved external-link state used while reading a workbook.
     * @throws FormatException    Thrown if {@code workbook} is null.
     * @throws FormatException    Thrown if {@code name} is null, empty, only whitespaces, starts with a digit or
     *                            contains illegal characters.
     * @throws FormatException    Thrown if {@code reference} is null or resolved to an empty string
     * @throws WorksheetException Thrown if {@code reference} already exists or matches a cell reference in the same
     *                            scope
     */
    DefinedName(
            Workbook workbook, NameType type, String name, Object reference, Worksheet worksheet, Worksheet localSheet,
            String comment, Optional<Boolean> containsExternalLinks
    ) {
        if (workbook == null) {
            throw new FormatException("To set a defined name, a workbook must be provided.");
        }
        validateName(workbook, name, localSheet);
        if (reference == null || ParserUtils.isNullOrEmpty(reference.toString())) // reference should never be null
        {
            throw new FormatException("The reference of a defined name must not be null or empty.");
        }
        this.type = type;
        this.name = name;
        this.value = reference;
        this.targetWorksheet = worksheet;
        this.localSheet = localSheet;
        this.comment = comment;
        castValue(workbook);
        this.features = FeatureSet.createDefinedName();
        this.features.add(workbook.getFeatures()); // Add feature reference already here
        if (containsExternalLinks.isPresent()) {

        }
        boolean hasExternalLinks = containsExternalLinks.orElseGet(
                () -> this.type == NameType.FORMULA
                        && ParserUtils.containsExternalReference(this.textValue)
        );
        boolean isFormula = this.type == NameType.FORMULA || hasExternalLinks;
        this.features.setDefinedNameFeatures(isFormula, hasExternalLinks);
    }
//methods

    /**
     * Resolve a defined name from its string reference
     *
     * @param name       Name (cannot be null)
     * @param reference  String reference (cannot be null)
     * @param workbook   Workbook reference (for cells, ranges and worksheet resolution)
     * @param localSheet Local sheet (can be null)
     * @param comment    Comment (can be null)
     * @return Resolved defined name object
     */
    static DefinedName resolveDefinedName(
            String name, String reference, Workbook workbook, Worksheet localSheet, String comment) {
        ParsedObject parsedObject = getParsedObject(reference);
        boolean optionalLinks = containsExternalLink(
                parsedObject.getWorksheet(), parsedObject.getType(), parsedObject.getValue());
        Worksheet worksheet = null;
        if (parsedObject.getWorksheet() != null && !optionalLinks) {
            for (Worksheet ws : workbook.getWorksheets()) {
                if (ParserUtils.equalsIgnoreCase(parsedObject.getWorksheet(), ws.getSheetName())) {
                    worksheet = ws;
                    break;
                }
            }
        }
        DefinedName definedName = new DefinedName(
                workbook, parsedObject.getType(), name, parsedObject.getValue(),
                worksheet, localSheet, comment, Optional.of(optionalLinks)
        );
        definedName.error = parsedObject.getError();
        return definedName;
    }

    /**
     * Internal Method to replace a string expression of a {@link NameType#FORMULA} instance. Other types are ignored
     *
     * <p>Remarks: This method is mainly supposed to replace external link tokes by compatibility plug-ins</p>
     *
     * @param expression New string expression
     */
    // TODO check whether it is possible to keep this non-public
    void ReplaceExpression(String expression) // Do not remove. This method may be used by NanoXLSX.Compatibility
    {
        if (type == NameType.FORMULA) {
            value = expression;
            textValue = expression;
        }
    }

    /**
     * Validates that the supplied name is non-empty, does not start with a digit, an invalid name or character, and
     * does not match a cell reference in the range A1 - XFD1048576.
     *
     * @param name       Name to validate.
     * @param workbook   Workbook to check for duplicate names
     * @param localSheet Local worksheet reference. Can be null if workbook scope
     * @throws FormatException Thrown if validation fails.
     */
    private static void validateName(Workbook workbook, String name, Worksheet localSheet) {
        if (ParserUtils.isNullOrWhiteSpace(name)) {
            throw new FormatException("The name of a defined name must not be null or empty.");
        }
        if (name.length() > MAX_NAME_LENGTH) {
            throw new FormatException("A defined name must not exceed " + MAX_NAME_LENGTH + " characters.");
        }

        char firstChar = name.charAt(0);

        if (!Character.isLetter(firstChar) && !contains(ALLOWED_NAME_START_CHARS, firstChar)) {
            throw new FormatException(
                    "The name of a defined name must start with a letter, underscore, or backslash. Provided: '" +
                            name + "'");
        }
        if (DISALLOWED_NAMES.contains(name)) {
            throw new FormatException("'" + name + "' cannot be used as a defined name.");
        }
        for (int i = 1; i < name.length(); i++) {
            char character = name.charAt(i);

            if (!Character.isLetterOrDigit(character) && !contains(ALLOWED_NAME_CHARS, character)) {
                throw new FormatException(
                        "The character '{character}' at position {i} is not valid in the defined name '" + name + "'.");
            }
        }
        if (workbook.findDefinedNameIndex(name, localSheet) >= 0) {
            String scope = localSheet == null ? "workbook" : "worksheet '" + localSheet.getSheetName() + "'";
            throw new WorksheetException(
                    "A defined name with the name '" + name + "' already exists in the " + scope + " scope.");
        }
        try {
            Validators.validateCellAddressExpression(name, Cell.AddressScope.SINGLE_ADDRESS);
        } catch (Exception e) {
            // Not a valid cell address; therefore it may be used as a defined name.
            return;
        }
        throw new FormatException("The defined name '" + name + "' must not be a valid cell address.");
    }

    /**
     * Casts {@link DefinedName#getValue()}  to a valid string for {@link DefinedName#getTextValue()}
     *
     * @param workbook Workbook instance
     * @throws FormatException Thrown if an expected address or range expression is invalid
     */
    private void castValue(Workbook workbook) {
        switch (this.type) {
            // The object type is assumed to be validated prior
            case NameType.CELL:
                String address =
                        this.value instanceof Address addressValue ? addressValue.toString() : this.value.toString();
                Validators.validateCellAddressExpression(
                        address, Cell.AddressScope.SINGLE_ADDRESS); // throw if not an address
                Address fixedAddress = new Address(address, Cell.AddressType.FIXED_ROW_AND_COLUMN);
                this.textValue = fixedAddress.toString();
                this.value = fixedAddress; // Reformat passed object
                break;
            case NameType.RANGE:
                String range = this.value instanceof Range rangeValue ? rangeValue.toString() : this.value.toString();
                Validators.validateCellAddressExpression(range, Cell.AddressScope.RANGE); // throw if not valid range
                Range tempRange = new Range(range);
                Range fixedRange = new Range(
                        new Address(
                                tempRange.startAddress().column(), tempRange.startAddress().row(),
                                Cell.AddressType.FIXED_ROW_AND_COLUMN
                        ), new Address(
                        tempRange.endAddress().column(), tempRange.endAddress().row(),
                        Cell.AddressType.FIXED_ROW_AND_COLUMN
                )
                );
                this.textValue = fixedRange.toString();
                this.value = fixedRange;
                break;
            case NameType.FORMULA:
                this.textValue = this.value.toString(); // No formula validation yet
                break;
            default: // constant
                this.textValue = ParserUtils.toCachedValueString(this.value, false);
                break;
        }
    }

    /**
     * Checks whether the defined name contains an external link (not applicable for constants)
     *
     * @param worksheet optional worksheet name
     * @param type      Type of the defined name
     * @param value     Value of the defined name
     * @return Returns true if an external link was found
     */
    private static boolean containsExternalLink(String worksheet, NameType type, Object value) {
        switch (type) {
            case NameType.FORMULA:
                String formula = (String) value;
                return ParserUtils.containsExternalReference(formula);
            case NameType.RANGE:
            case NameType.CELL:
                return worksheet != null && EXT_WORKSHEET_REFERENCE_REGEX.matcher(worksheet).matches();
            default: // constant
                break; // NoOp
        }
        return false;
    }

    /**
     * Gets the parsed object from a raw reference string.
     *
     * @param reference Raw reference string to parse
     * @return Parsed value and its associated defined-name metadata
     */
    private static ParsedObject getParsedObject(String reference) {
        ParsedObject result = new ParsedObject();
        if (reference != null) {
            reference = reference.trim();
        }
        result.setError(FormulaError.NO_ERROR); // Initial state
        result.setWorksheet(null); // Initial state

        Optional<String> stringValue = ParserUtils.tryParseFormulaStringConstant(reference);
        if (stringValue.isPresent()) {
            result.setType(NameType.CONSTANT); // Formula string is interpreted as constant in this case
            result.setValue(stringValue.get());
            return result;
        }

        Optional<Boolean> booleanValue = ParserUtils.tryParseBoolean(reference);
        if (booleanValue.isPresent()) {
            result.setType(NameType.CONSTANT);
            result.setValue(booleanValue.get());
            return result;
        }

        Optional<Integer> intValue = ParserUtils.tryParseInt(reference);
        if (intValue.isPresent()) {
            result.setType(NameType.CONSTANT);
            result.setValue(intValue.get());
            return result;
        }

        Optional<Double> doubleValue = ParserUtils.tryParseDouble(reference);
        if (doubleValue.isPresent()) {
            result.setType(NameType.CONSTANT);
            result.setValue(doubleValue.get());
            return result;
        }

        Optional<ParserUtils.WorksheetQualifiedReference> qualifiedReference =
                ParserUtils.tryParseWorksheetQualifiedReference(reference);
        if (qualifiedReference.isPresent()) {
            ParserUtils.WorksheetQualifiedReference parsedReference = qualifiedReference.get();
            result.setWorksheet(parsedReference.worksheetName());
            try {
                Address addressValue = new Address(parsedReference.reference());
                result.setType(NameType.CELL);
                result.setValue(addressValue);
                return result;
            } catch (Exception e) {
                // NoOp
            }
            try {
                Range range = new Range(parsedReference.reference());
                result.setType(NameType.RANGE);
                result.setValue(range);
                return result;
            } catch (Exception e) {
                // NoOp
            }
        }

        Optional<FormulaError> errorRef = Errors.tryParseFormulaError(reference);
        errorRef.ifPresent(result::setError);
        result.setType(NameType.FORMULA);
        result.setValue(reference);
        return result;
    }

    /**
     * Compares this instance with another {@link DefinedName} for ordering. The order is determined by
     * {@link DefinedName#getName()} (ordinal), then by scope (workbook scope sorts before any worksheet-scoped name;
     * for worksheet-scoped names by {@link Worksheet#getSheetId()}), then by {@link DefinedName#getTextValue()}, then
     * by {@link  DefinedName#getComment()}.
     *
     * @param o Other defined name, or null. A null comparand sorts after this instance.
     * @return Negative, zero, or positive integer following the standard {@link Comparator{T}} contract.
     */
    @Override
    public int compareTo(DefinedName o) {
        if (o == null) {
            return 1;
        }
        int cmp = Comparator.nullsFirst(String::compareTo).compare(name, o.name);
        if (cmp != 0) {
            return cmp;
        }
        cmp = type.compareTo(o.type);
        if (cmp != 0) {
            return cmp;
        }
        cmp = compareScope(localSheet, o.localSheet);
        if (cmp != 0) {
            return cmp;
        }
        cmp = compareScope(targetWorksheet, o.targetWorksheet);
        if (cmp != 0) {
            return cmp;
        }
        cmp = Comparator.nullsFirst(String::compareTo).compare(textValue, o.textValue);
        if (cmp != 0) {
            return cmp;
        }
        return Comparator.nullsFirst(String::compareTo).compare(comment, o.comment);
    }

    /**
     * Compares two scope worksheets for ordering. Workbook scope (null) sorts before any worksheet scope; two worksheet
     * scopes are ordered by their {@link Worksheet#getSheetId()}.
     *
     * @param left  Left scope, or null for workbook scope.
     * @param right Right scope, or null for workbook scope.
     * @return Negative, zero, or positive integer following the standard ordering contract.
     */
    private static int compareScope(Worksheet left, Worksheet right) {
        if (left == right) {
            return 0;
        }
        if (left == null) {
            return -1;
        }
        if (right == null) {
            return 1;
        }
        return Comparator.nullsFirst(Integer::compareTo).compare(left.getSheetId(), right.getSheetId());
    }

    /**
     * Checks whether a char array contains a specific char
     *
     * @param values Char array
     * @param target Car to check
     * @return True, if the char is present, otherwise false
     */
    private static boolean contains(char[] values, char target) {
        for (char value : values) {
            if (value == target) {
                return true;
            }
        }
        return false;
    }

    /**
     * Helper class for parsed objects and all relevant parameters
     */
    private static class ParsedObject {
        private Object value;
        private NameType type;
        private String worksheet;
        private FormulaError error;

        /**
         * Gets the object value
         *
         * @return Object value
         */
        public Object getValue() {
            return value;
        }

        /**
         * Sets the object value
         *
         * @param value Object value
         */
        public void setValue(Object value) {
            this.value = value;
        }

        /**
         * Gets the name type of the object
         *
         * @return Object type
         */
        public NameType getType() {
            return type;
        }

        /**
         * Sets the name type of the object
         *
         * @param type Object type
         */
        public void setType(NameType type) {
            this.type = type;
        }

        /**
         * Gets the worksheet name
         *
         * @return Worksheet name
         */
        public String getWorksheet() {
            return worksheet;
        }

        /**
         * Sets the worksheet name
         *
         * @param worksheet Worksheet name
         */
        public void setWorksheet(String worksheet) {
            this.worksheet = worksheet;
        }

        /**
         * Gets the formula error if applicable
         *
         * @return Formula error
         */
        public FormulaError getError() {
            return error;
        }

        /**
         * Sets the formula error if applicable
         *
         * @param error Formula error
         */
        public void setError(FormulaError error) {
            this.error = error;
        }
    }

}
