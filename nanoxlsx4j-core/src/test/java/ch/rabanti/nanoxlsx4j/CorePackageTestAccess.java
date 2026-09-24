package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;

/** Access to package-private Core methods for tests that mirror the C# directory structure. */
public final class CorePackageTestAccess {

    private CorePackageTestAccess() {
    }

    public static void setNullReference(Cell cell) {
        cell.setReference(null);
    }

    public static void setFormula(Cell cell, FormulaData formula) {
        cell.setFormula(formula);
    }

    public static void setCachedValueType(FormulaData formula, Cell.CellType type) {
        formula.setCachedValueType(type);
    }

    public static void setFormulaExpression(FormulaData formula, String expression) {
        formula.setExpression(expression);
    }

    public static void setFormulaType(FormulaData formula, FormulaData.FormulaType type) {
        formula.setType(type);
    }

    public static void setFormulaRange(FormulaData formula, String range) {
        formula.setFormulaRange(range);
    }

    public static FormulaData copyFormulaData(FormulaData formula) {
        return formula.copy();
    }

    public static boolean hasSameFeatureSet(FormulaData first, FormulaData second) {
        return first.getFeatures() == second.getFeatures();
    }

    public static Cell copyCell(Cell cell) {
        return cell.copy();
    }

    public static int worksheetFormulaCount(Worksheet worksheet) {
        return worksheet.getFeatures().getFormulaCount();
    }

    public static int worksheetExternalLinkCount(Worksheet worksheet) {
        return worksheet.getFeatures().getExternalLinkCount();
    }

    public static void recalculateAutoFilter(Worksheet worksheet) {
        worksheet.recalculateAutoFilter();
    }

    public static void recalculateColumns(Worksheet worksheet) {
        worksheet.recalculateColumns();
    }

    public static void resolveMergedCells(Worksheet worksheet) {
        worksheet.resolveMergedCells();
    }

    public static void setSheetProtectionPassword(Worksheet worksheet, Password password) {
        worksheet.setSheetProtectionPassword(password);
    }
}
