package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.internal.FeatureSet;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotSame;

public class FeatureSetTest {

    @Test
    @DisplayName("A new aggregate feature set has no features")
    public void constructorCreatesEmptyAggregateFeatureSet() {
        FeatureSet featureSet = new FeatureSet();

        assertFeatures(featureSet, 0, 0, 0, 0);
    }

    @Test
    @DisplayName("The formula factory creates a valid formula feature set")
    public void createFormulaCreatesFormulaFeatureSet() {
        FeatureSet featureSet = FeatureSet.createFormula();

        assertFeatures(featureSet, 1, 0, 0, 0);
    }

    @Test
    @DisplayName("The defined-name factory creates a valid defined-name feature set")
    public void createDefinedNameCreatesDefinedNameFeatureSet() {
        FeatureSet featureSet = FeatureSet.createDefinedName();

        assertFeatures(featureSet, 0, 1, 0, 0);
    }

    @Test
    @DisplayName("Adding and removing feature sets updates all counts through the parent hierarchy")
    public void addAndRemovePropagatesAllFeatures() {
        FeatureSet root = new FeatureSet();
        FeatureSet parent = new FeatureSet();
        FeatureSet aggregate = new FeatureSet();
        FeatureSet formula = FeatureSet.createFormula();
        FeatureSet definedName = FeatureSet.createDefinedName();

        parent.add(root);
        formula.setFormulaFeatures(true, true);
        definedName.setDefinedNameFeatures(true, true);
        formula.add(aggregate);
        definedName.add(aggregate);

        assertFeatures(aggregate, 2, 1, 1, 2, 1);
        assertFeatures(parent, 0, 0, 0, 0);
        assertFeatures(root, 0, 0, 0, 0);

        aggregate.add(parent);

        assertFeatures(parent, 2, 1, 1, 2, 1);
        assertFeatures(root, 2, 1, 1, 2, 1);

        formula.Remove(aggregate);

        assertFeatures(formula, 1, 0, 0, 1, 1);
        assertFeatures(aggregate, 1, 1, 1, 1);
        assertFeatures(parent, 1, 1, 1, 1);
        assertFeatures(root, 1, 1, 1, 1);

        definedName.Remove(aggregate);

        assertFeatures(definedName, 1, 1, 1, 1);
        assertFeatures(aggregate, 0, 0, 0, 0);
        assertFeatures(parent, 0, 0, 0, 0);
        assertFeatures(root, 0, 0, 0, 0);
    }

    @ParameterizedTest
    @DisplayName("Setting formula features updates local and parent feature values")
    @CsvSource({
        "false, false",
        "false, true",
        "true, false",
        "true, true"
    })
    public void setFormulaFeaturesUpdatesFeatureValues(
            boolean containsDefinedName,
            boolean containsExternalLink) {
        FeatureSet root = new FeatureSet();
        FeatureSet parent = new FeatureSet();
        FeatureSet formula = FeatureSet.createFormula();
        parent.add(root);
        formula.add(parent);

        // Start from the opposite state so every InlineData case exercises a transition.
        formula.setFormulaFeatures(!containsDefinedName, !containsExternalLink);
        // Test transitions from false to true and from true to false.
        formula.setFormulaFeatures(containsDefinedName, containsExternalLink);

        int definedNameReferenceFormulaCount = containsDefinedName ? 1 : 0;
        int externalLinkCount = containsExternalLink ? 1 : 0;
        assertFeatures(formula, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);
        assertFeatures(parent, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);
        assertFeatures(root, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);

        formula.setFormulaFeatures(containsDefinedName, containsExternalLink);

        assertFeatures(formula, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);
        assertFeatures(parent, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);
        assertFeatures(root, 1, 0, 0, externalLinkCount, definedNameReferenceFormulaCount);
    }

    @ParameterizedTest
    @DisplayName("Setting defined-name features updates local and parent feature values")
    @CsvSource({
        "false, false",
        "false, true",
        "true, false",
        "true, true"
    })
    public void setDefinedNameFeaturesUpdatesFeatureValues(boolean isFormula, boolean containsExternalLink) {
        FeatureSet root = new FeatureSet();
        FeatureSet parent = new FeatureSet();
        FeatureSet definedName = FeatureSet.createDefinedName();
        parent.add(root);
        definedName.add(parent);

        // Start from the opposite state so every InlineData case exercises a transition.
        definedName.setDefinedNameFeatures(!isFormula, !containsExternalLink);
        // Test transitions from false to true and from true to false.
        definedName.setDefinedNameFeatures(isFormula, containsExternalLink);

        int formulaCount = isFormula ? 1 : 0;
        int externalLinkCount = containsExternalLink ? 1 : 0;
        assertFeatures(definedName, formulaCount, 1, formulaCount, externalLinkCount);
        assertFeatures(parent, formulaCount, 1, formulaCount, externalLinkCount);
        assertFeatures(root, formulaCount, 1, formulaCount, externalLinkCount);

        definedName.setDefinedNameFeatures(isFormula, containsExternalLink);

        assertFeatures(definedName, formulaCount, 1, formulaCount, externalLinkCount);
        assertFeatures(parent, formulaCount, 1, formulaCount, externalLinkCount);
        assertFeatures(root, formulaCount, 1, formulaCount, externalLinkCount);
    }

    @Disabled("TODO: Requires Workbook and defined-name support")
    @Test
    @DisplayName("A defined-name formula with an external link is not counted as a worksheet formula")
    public void definedNameExternalFormulaUpdatesWorkbookFeatures() {
        // TODO: C# reference: NanoXlsx.Core.Test/FeatureSetTest.cs; missing Workbook defined-name support.
    }

    @Test
    @DisplayName("Copying a feature set preserves its values without retaining its parent")
    public void copyPreservesValuesWithoutParent() {
        FeatureSet parent = new FeatureSet();
        FeatureSet original = FeatureSet.createFormula();
        original.setFormulaFeatures(true, true);
        original.add(parent);

        FeatureSet copy = original.copy();

        assertNotSame(original, copy);
        assertFeatures(copy, 1, 0, 0, 1, 1);

        copy.setFormulaFeatures(false, false);

        assertFeatures(copy, 1, 0, 0, 0);
        assertFeatures(original, 1, 0, 0, 1, 1);
        assertFeatures(parent, 1, 0, 0, 1, 1);
    }

    @Disabled("TODO: Requires Workbook and the complete Cell.CellType model")
    @Test
    @DisplayName("Changing a formula cell to a non-formula type removes its feature contribution")
    public void cellFormulaTypeChangeRemovesFeatures() {
        // TODO: C# reference: NanoXlsx.Core.Test/FeatureSetTest.cs; missing Workbook and Cell.CellType.
    }

    @Disabled("TODO: Requires Workbook and formula metadata support")
    @Test
    @DisplayName("A cached formula error preserves formula and external-link features")
    public void cellFormulaCachedErrorPreservesFeatures() {
        // TODO: C# reference: NanoXlsx.Core.Test/FeatureSetTest.cs; missing Workbook and FormulaData.
    }

    @Disabled("TODO: Requires Workbook and FormulaData")
    @Test
    @DisplayName("Replacing formula metadata keeps feature counters balanced")
    public void cellFormulaReplacementUpdatesFeatures() {
        // TODO: C# reference: NanoXlsx.Core.Test/FeatureSetTest.cs; missing Workbook and FormulaData.
    }

    @Disabled("TODO: Requires Workbook, DefinedName, and formula metadata support")
    @Test
    @DisplayName("Changing a defined-name formula value removes its resolved reference feature")
    public void cellDefinedNameFormulaValueChangeUpdatesFeatures() {
        // TODO: C# reference: NanoXlsx.Core.Test/FeatureSetTest.cs; missing Workbook, DefinedName, and FormulaData.
    }

    private static void assertFeatures(
            FeatureSet featureSet,
            int formulaCount,
            int definedNameCount,
            int definedNameFormulaCount,
            int externalLinkCount) {
        assertFeatures(featureSet, formulaCount, definedNameCount, definedNameFormulaCount, externalLinkCount, 0);
    }

    private static void assertFeatures(
            FeatureSet featureSet,
            int formulaCount,
            int definedNameCount,
            int definedNameFormulaCount,
            int externalLinkCount,
            int definedNameReferenceFormulaCount) {
        assertEquals(formulaCount, featureSet.getFormulaCount());
        assertEquals(definedNameCount, featureSet.getDefinedNameCount());
        assertEquals(definedNameFormulaCount, featureSet.getDefinedNameFormulaCount());
        assertEquals(definedNameReferenceFormulaCount, featureSet.getDefinedNameReferenceFormulaCount());
        assertEquals(formulaCount - definedNameFormulaCount, featureSet.getWorksheetFormulaCount());
        assertEquals(externalLinkCount, featureSet.getExternalLinkCount());
        assertEquals(formulaCount > 0, featureSet.containsFormulas());
        assertEquals(definedNameCount > 0, featureSet.containsDefinedName());
        assertEquals(definedNameFormulaCount > 0, featureSet.containsDefinedNameFormula());
        assertEquals(definedNameReferenceFormulaCount > 0, featureSet.containsDefinedNameReferenceFormula());
        assertEquals(formulaCount - definedNameFormulaCount > 0, featureSet.containsWorksheetFormula());
        assertEquals(externalLinkCount > 0, featureSet.containsExternalLink());
    }
}
