package ch.rabanti.nanoxlsx4j;

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

    public static Cell copyCell(Cell cell) {
        return cell.copy();
    }

    public static int worksheetFormulaCount(Worksheet worksheet) {
        return worksheet.getFeatures().getFormulaCount();
    }

    public static int worksheetExternalLinkCount(Worksheet worksheet) {
        return worksheet.getFeatures().getExternalLinkCount();
    }
}
