/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import java.math.BigDecimal;
import java.time.Duration;
import java.util.Comparator;
import java.util.Date;
import java.util.Objects;

import ch.rabanti.nanoxlsx4j.enums.FormulaError;
import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

/**
 * Class representing a formula in a cell, its data, respectively
 */
public class FormulaData implements Comparable<FormulaData> {

    /**
     * Enum to define the specific type of a formula if the Cell has the type {@link Cell.CellType#FORMULA}
     */
    public enum FormulaType {
        /**
         * Cell contains a regular formula (e.g "A1+A2")
         */
        NORMAL,
        /**
         * Cell contains a formula that is part of an array
         */
        ARRAY,
        /**
         * Cell contains a shared formula, pointing to another formula that is identical
         */
        SHARED,
        /**
         * Cell contains a formula that is applied across a range of one or more cells
         */
        DATA_TABLE
    }

    private DefinedName definedNameReference;
    private String expression;
    private FormulaType type;
    private Object cachedValue;
    private Cell.CellType cachedValueType;
    private String masterCellAddress;
    private FeatureSet features;
    private boolean externalReferences;
    private String formulaRange;

    /**
     * Gets the formula expression as string. This value is currently identical with {@link Cell#getValue()} if
     * {@link Cell#getDataType()} is set to {@link Cell.CellType#FORMULA}.
     *
     * @return Formula expression as string
     */
    public String getExpression() {
        return expression;
    }

    /**
     * Sets the formula expression internally
     *
     * @param expression Formula expression as string
     */
    void setExpression(String expression) {
        this.expression = expression;
        setExternalReferences(ParserUtils.containsExternalReference(expression));
    }

    public DefinedName getDefinedNameReference() {
        return definedNameReference;
    }

    public void setDefinedNameReference(DefinedName definedNameReference) {
        this.definedNameReference = definedNameReference;
    }

    /**
     * Gets the type of the formula. Default is {@link FormulaType#NORMAL}
     *
     * @return Formula type
     */
    public FormulaType getType() {
        return type;
    }

    /**
     * Sets the type of the formula internally
     *
     * @param formulaType Formula type
     */
    void setType(FormulaType formulaType) {
        this.type = formulaType;
    }

    /**
     * Gets the range associated with an array, shared, or data-table formula. Can be a range or address (string
     * representation). Default is null, if no reference was defined.
     *
     * @return Formula range (can be null)
     */
    public String getFormulaRange() {
        return formulaRange;
    }

    /**
     * Sets the formula range internally
     *
     * @param formulaRange Formula range (can be null)
     */
    void setFormulaRange(String formulaRange) {
        this.formulaRange = formulaRange;
    }

    /**
     * Gets whether the formula expression contains a reference to an external workbook.
     *
     * @return If true, the formula contains external references
     */
    public boolean hasExternalReferences() {
        return externalReferences;
    }

    /**
     * Sets whether the formula contains external references (internal setter)
     *
     * @param externalReferences If true, the formula contains external references
     */
    void setExternalReferences(boolean externalReferences) {
        this.externalReferences = externalReferences;
        boolean isDefinedName = definedNameReference != null;
        this.features.setFormulaFeatures(isDefinedName, externalReferences);
    }

    /**
     * Gets the cached value of the formula
     *
     * <p>Remarks: This value can be supplied through the constructor or set when a Workbook is loaded. Formula error
     * results are represented by {@link FormulaError} values. The value is not evaluated when a new formula was defined
     * by {@link Worksheet#addCellFormula(String, int, int)} or its overload methods.
     * </p>
     *
     * @return Cached formula value
     */
    public Object getCachedValue() {
        return cachedValue;
    }

    /**
     * Sets the cached value of the formula internally
     *
     * @param textValue Cached formula value
     */
    public void setCachedValue(Object textValue) {
        this.cachedValue = textValue;
    }

    /**
     * Gets the data type of {@link FormulaData#getCachedValue()}. The default is {@link Cell.CellType#DEFAULT} if no
     * cached value or no supported cached value type is available.
     *
     * @return Type of the cached value
     */
    public Cell.CellType getCachedValueType() {
        return cachedValueType;
    }

    /**
     * Sets the type of the cached value internally
     *
     * @param cachedValueType Type of the cached value
     */
    void setCachedValueType(Cell.CellType cachedValueType) {
        this.cachedValueType = cachedValueType;
    }

    /**
     * Gets the address of the formula's master cell. This is mainly used in case of {@link FormulaType#ARRAY}.
     *
     * @return Master cell address as string (can be null)
     */
    public String getMasterCellAddress() {
        return masterCellAddress;
    }

    /**
     * Sets the master cell address internally
     *
     * @param masterCellAddress Master cell address as string (can be null)
     */
    public void setMasterCellAddress(String masterCellAddress) {
        this.masterCellAddress = masterCellAddress;
    }

    /**
     * Internal feature set for cascading feature detection (consider in {@link FormulaData#copy()}  but not in Equals,
     * GetHashCode etc.)
     *
     * @return Feature set
     */
    FeatureSet getFeatures() {
        return features;
    }

    /**
     * Sets the feature set of the formula internally
     *
     * @param features Feature set
     */
    void setFeatures(FeatureSet features) {
        this.features = features;
    }

    /**
     * Default constructor
     */
    public FormulaData() {
        this.type = FormulaType.NORMAL;
        this.cachedValueType = Cell.CellType.DEFAULT;
        this.features = FeatureSet.createFormula();
    }

    /**
     * Constructor with formula expression and optional cached value to create a formula of the common type
     * {@link FormulaType#NORMAL}
     *
     * <p>Remarks: A basic validity checks (not full parsing) will perform on the expression, e.g. existence of an
     * external link in the formula</p>
     *
     * @param expression Formula expression (without leading equal sign)
     */
    public FormulaData(String expression) {
        this(expression, null);
    }

    /**
     * Constructor with formula expression and optional cached value to create a formula of the common type
     * {@link FormulaType#NORMAL}
     *
     * <p>Remarks: A basic validity checks (not full parsing) will perform on the expression, e.g. existence of an
     * external link in the formula</p>
     *
     * @param expression  Formula expression (without leading equal sign)
     * @param cachedValue Optional cached value. Default is null
     */
    public FormulaData(String expression, Object cachedValue) {
        this();
        setExpression(expression);
        this.cachedValue = cachedValue;
        cachedValueType = resolveCachedValueType(cachedValue);
    }

    /**
     * Resolves the cell type of a cached formula value without evaluating the formula.
     *
     * @param cachedValue Cached formula value, or null if unavailable.
     * @return Resolved cached value type.
     */
    static Cell.CellType resolveCachedValueType(Object cachedValue) {
        if (cachedValue == null) {
            return Cell.CellType.DEFAULT;
        }
        if (cachedValue instanceof Boolean) {
            return Cell.CellType.BOOL;
        }
        if (cachedValue instanceof Byte || cachedValue instanceof BigDecimal || cachedValue instanceof Double
                || cachedValue instanceof Float || cachedValue instanceof Integer || cachedValue instanceof Long
                || cachedValue instanceof Short) {
            return Cell.CellType.NUMBER;
        }
        if (cachedValue instanceof Date) {
            return Cell.CellType.DATE;
        }
        if (cachedValue instanceof Duration) {
            return Cell.CellType.TIME;
        }
        if (cachedValue instanceof FormulaError) {
            return Cell.CellType.ERROR;
        }
        return Cell.CellType.STRING;
    }

    /**
     * Copies the current object into a new one (without copying {@link FormulaData#getDefinedNameReference()})
     *
     * <p>Remarks: This copy method omits deep-copying {@link FormulaData#getDefinedNameReference()} by design. If a
     * full copy is intended, this instance variable must be handled separately.</p>
     *
     * @return Copy of the current instance
     */
    FormulaData copy() {
        FormulaData data = new FormulaData(expression); // For automatic feature resolution
        data.type = this.type;
        data.formulaRange = this.formulaRange;
        data.cachedValue = this.cachedValue;
        data.cachedValueType = this.cachedValueType;
        data.masterCellAddress = this.masterCellAddress;
        data.features = this.features.copy(); // New feature set
        data.definedNameReference = this.definedNameReference; // object reference
        return data;
    }

    /**
     * Compares this instance with another {@link FormulaData} instance.
     *
     * @param o Other formula data instance, or null.
     * @return Negative, zero, or positive integer following the standard comparison contract.
     */
    @Override
    public int compareTo(FormulaData o) {
        if (o == null) {
            return 1;
        }
        int cmp = Comparator.nullsFirst(String::compareTo).compare(expression, o.expression);
        if (cmp != 0) {
            return cmp;
        }
        cmp = type.compareTo(o.type);
        if (cmp != 0) {
            return cmp;
        }
        cmp = Comparator.nullsFirst(String::compareTo).compare(formulaRange, o.formulaRange);
        if (cmp != 0) {
            return cmp;
        }
        cmp = Comparator.nullsFirst(DefinedName::compareTo).compare(definedNameReference, o.definedNameReference);
        if (cmp != 0) {
            return cmp;
        }
        cmp = cachedValueType.compareTo(o.cachedValueType);
        if (cmp != 0) {
            return cmp;
        }
        cmp = compareObjects(cachedValue, o.cachedValue);
        if (cmp != 0) {
            return cmp;
        }
        return Comparator.nullsFirst(String::compareTo).compare(masterCellAddress, o.masterCellAddress);
    }

    /**
     * Determines whether the specified object is equal to this instance.
     *
     * @param o the reference object with which to compare.
     * @return True if equal, otherwise false.
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof FormulaData that)) {
            return false;
        }

        return externalReferences == that.externalReferences &&
                Objects.equals(definedNameReference, that.definedNameReference) &&
                Objects.equals(expression, that.expression) && type == that.type &&
                Objects.equals(cachedValue, that.cachedValue) && cachedValueType == that.cachedValueType &&
                Objects.equals(masterCellAddress, that.masterCellAddress) &&
                Objects.equals(formulaRange, that.formulaRange);
    }

    /**
     * Returns a hash code consistent with {@link FormulaData#equals(Object)}.
     *
     * @return Hash code derived from this instance's properties.
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(definedNameReference);
        result = 31 * result + Objects.hashCode(expression);
        result = 31 * result + Objects.hashCode(type);
        result = 31 * result + Objects.hashCode(cachedValue);
        result = 31 * result + Objects.hashCode(cachedValueType);
        result = 31 * result + Objects.hashCode(masterCellAddress);
        result = 31 * result + Boolean.hashCode(externalReferences);
        result = 31 * result + Objects.hashCode(formulaRange);
        return result;
    }

    /**
     * Compares two generic objects for equality
     *
     * @param a Object 1
     * @param b Object 2
     * @return Negative, zero, or positive integer following the standard comparison contract
     */
    @SuppressWarnings(
            {
                    "rawtypes",
                    "unchecked"}
    )
    private static int compareObjects(Object a, Object b) {
        if (a == b) {
            return 0;
        }
        if (a == null) {
            return -1;
        }
        if (b == null) {
            return 1;
        }
        if (a instanceof Comparable comparable) {
            return comparable.compareTo(b);
        }
        throw new ClassCastException("Objects are not comparable: " + a.getClass().getName());
    }
}
