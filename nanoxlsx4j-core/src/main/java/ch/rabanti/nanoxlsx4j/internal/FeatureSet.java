/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.internal;

/**
 * Class to count and aggregate features on the level of cells, worksheets and workbooks.
 * <p>Remarks: Internal use only, without checks. Do not tamper with the class or instances</p>
 */
public final class FeatureSet {

    /**
     * Enum of the type of the feature set
     */
    private enum FeatureSetType {
        /** Feature set is between the root and leaf element */
        AGGREGATE,
        /** Feature set is on a formula (leaf element) */
        FORMULA,
        /** Feature set is on defined name instance (leaf element) */
        DEFINED_NAME
    }

    private FeatureSetType type;
    private FeatureSet parent;
    private int formulaCount;
    private int definedNameCount;
    private int definedNameFormulaCount;
    private int definedNameReferenceFormulaCount;
    private int externalLinkCount;

    /**
     * Gets the number of formulas in the feature set.
     *
     * @return Feature count
     */
    public int getFormulaCount() {
        return formulaCount;
    }

    /**
     * gets the Number of defined-name definitions in the feature set.
     *
     * @return Feature count
     */
    public int getDefinedNameCount() {
        return definedNameCount;
    }

    /**
     * Gets the number of formulas that are defined by defined names rather than worksheet cells.
     *
     * @return Feature count
     */
    public int getDefinedNameFormulaCount() {
        return definedNameFormulaCount;
    }

    /**
     * Gets the number of worksheet formulas that directly reference a resolved defined name.
     * <p>Remarks: Defined names embedded within larger formula expressions are currently not resolved and are not
     * counted.</p>
     *
     * @return Feature count
     */
    public int getDefinedNameReferenceFormulaCount() {
        return definedNameReferenceFormulaCount;
    }

    /**
     * Gets the number of formulas that are on worksheets and their cells, but not on defined names. This may only be
     * useful on {@link FeatureSetType#AGGREGATE}
     *
     * @return Feature count
     */
    public int getWorksheetFormulaCount() {
        return formulaCount - definedNameFormulaCount;
    }

    /**
     * Gets the number external links in the feature set
     *
     * @return Feature count
     */
    public int getExternalLinkCount() {
        return externalLinkCount;
    }

    /**
     * Determines whether the feature set contains formulas (defined names and cells)
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsFormulas() {
        return formulaCount > 0;
    }

    /**
     * Determines whether the feature set contains defined names
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsDefinedName() {
        return definedNameCount > 0;
    }

    /**
     * Determines whether the feature set contains formulas in defined names
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsDefinedNameFormula() {
        return definedNameFormulaCount > 0;
    }

    /**
     * Determines whether the feature set contains worksheet formulas that directly reference resolved defined names
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsDefinedNameReferenceFormula() {
        return definedNameReferenceFormulaCount > 0;
    }

    /**
     * Determines whether the feature set contains formulas on worksheets and their cells
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsWorksheetFormula() {
        return  getWorksheetFormulaCount() > 0;
    }

    /**
     * Determines whether the feature set contains external links (defined names and cells)
     *
     * @return True if the feature has a count of >0 in this set
     */
    public boolean containsExternalLink() {
        return externalLinkCount > 0;
    }

    /**
     * Creates a feature set representing exactly one formula.
     *
     * @return FeatureSet instance
     */
    public static FeatureSet createFormula() {
        return new FeatureSet(FeatureSetType.FORMULA);
    }

    /**
     * Creates a feature set representing exactly one defined name
     *
     * @return FeatureSet instance
     */
    public static FeatureSet createDefinedName() {
        return new FeatureSet(FeatureSetType.DEFINED_NAME);
    }

    /**
     * Create new scratch file from selection
     */
    public FeatureSet() {
        type = FeatureSetType.AGGREGATE;
    }

    /**
     * Creates a feature set representing a non-aggregate feature
     *
     * @param type Type of the feature set
     */
    private FeatureSet(FeatureSetType type) {
        this.type = type;
        if (type == FeatureSetType.FORMULA) {
            formulaCount = 1;
        } else if (type == FeatureSetType.DEFINED_NAME) {
            definedNameCount = 1;
        }
    }

    /**
     * Adds this feature set to the specified parent. The current counts are propagated to the complete parent
     * hierarchy.
     *
     * @param parent Parent feature set
     */
    public void add(FeatureSet parent) {
        parent.applyDelta(
                formulaCount,
                definedNameCount,
                definedNameFormulaCount,
                definedNameReferenceFormulaCount,
                externalLinkCount
        );

        this.parent = parent;
    }

    /**
     * Removes this feature set from the specified parent. The current counts are subtracted from the complete parent
     * hierarchy.
     *
     * @param parent Parent feature set
     */
    public void Remove(FeatureSet parent) {
        parent.applyDelta(
                -formulaCount,
                -definedNameCount,
                -definedNameFormulaCount,
                -definedNameReferenceFormulaCount,
                -externalLinkCount
        );

        this.parent = null;
    }

    /**
     * Updates the features of a single formula and propagates all changes to the parent hierarchy.
     *
     * @param containsDefinedName  If true, the feature for defined names will be set
     * @param containsExternalLink If true, the feature for external links will be set
     */
    public void setFormulaFeatures(boolean containsDefinedName, boolean containsExternalLink) {
        int newDefinedNameReferenceFormulaCount = containsDefinedName ? 1 : 0;
        int newExternalLinkCount = containsExternalLink ? 1 : 0;

        int definedNameReferenceFormulaDelta =
                newDefinedNameReferenceFormulaCount - definedNameReferenceFormulaCount;

        int externalLinkDelta =
                newExternalLinkCount - externalLinkCount;

        if (definedNameReferenceFormulaDelta == 0 && externalLinkDelta == 0) {
            return;
        }

        if (parent != null) {
            parent.applyDelta(
                    0,
                    0,
                    0,
                    definedNameReferenceFormulaDelta,
                    externalLinkDelta
            );
        }

        definedNameReferenceFormulaCount = newDefinedNameReferenceFormulaCount;
        externalLinkCount = newExternalLinkCount;
    }

    /**
     * Updates the features of a single defined name and propagates all changes to the parent hierarchy.
     *
     * @param isFormula            If true, the feature for formulas will be set
     * @param containsExternalLink If true, the feature for external links will be set
     */
    public void setDefinedNameFeatures(boolean isFormula, boolean containsExternalLink) {
        int newFormulaCount = isFormula ? 1 : 0;
        int formulaDelta = newFormulaCount - formulaCount;

        int newDefinedNameFormulaCount = isFormula ? 1 : 0;
        int definedNameFormulaDelta = newDefinedNameFormulaCount - definedNameFormulaCount;

        int newExternalLinkCount = containsExternalLink ? 1 : 0;
        int externalLinkDelta = newExternalLinkCount - externalLinkCount;

        if (externalLinkDelta == 0 && formulaDelta == 0 && definedNameFormulaDelta == 0) {
            return;
        }

        if (parent != null) {
            parent.applyDelta(formulaDelta, 0, definedNameFormulaDelta, 0, externalLinkDelta);
        }

        formulaCount = newFormulaCount;
        definedNameFormulaCount = newDefinedNameFormulaCount;
        externalLinkCount = newExternalLinkCount;
    }

    /**
     * Applies a count difference to this aggregate FeatureSet and recursively propagates it to its parent
     *
     * @param formulaDelta                     Delta value for formula count
     * @param definedNameDelta                 Delta value for defined name count
     * @param definedNameFormulaDelta          Delta value for formulas defined by defined names
     * @param definedNameReferenceFormulaDelta Delta value for worksheet formulas referencing defined names
     * @param externalLinkDelta                Delta value for external link count
     */
    private void applyDelta(
            int formulaDelta,
            int definedNameDelta,
            int definedNameFormulaDelta,
            int definedNameReferenceFormulaDelta,
            int externalLinkDelta
    ) {
        formulaCount += formulaDelta;
        definedNameCount += definedNameDelta;
        definedNameFormulaCount += definedNameFormulaDelta;
        definedNameReferenceFormulaCount += definedNameReferenceFormulaDelta;
        externalLinkCount += externalLinkDelta;

        if (parent != null) {
            parent.applyDelta(
                    formulaDelta,
                    definedNameDelta,
                    definedNameFormulaDelta,
                    definedNameReferenceFormulaDelta,
                    externalLinkDelta
            );
        }
    }

    /**
     * Copies the instance with counters bot not with parent. The parent will be added by regular Add functions (e.g.
     * {@link ch.rabanti.nanoxlsx4j.Worksheet#addNextCellFormula(String)})
     *
     * @return Returns the copied instance
     */
    public FeatureSet copy() {
        FeatureSet copy = new FeatureSet(type);
        copy.formulaCount = formulaCount;
        copy.definedNameCount = definedNameCount;
        copy.definedNameFormulaCount = definedNameFormulaCount;
        copy.definedNameReferenceFormulaCount = definedNameReferenceFormulaCount;
        copy.externalLinkCount = externalLinkCount;
        return copy;
    }

}
