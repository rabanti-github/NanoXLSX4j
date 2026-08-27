package ch.rabanti.nanoxlsx4j;

import java.util.Comparator;

public class DefinedName implements Comparable<DefinedName>{

    /**
     * Enum to specify the type of the defined name
     */
    public enum NameType
    {
        /** Defined name is a single cell */
        CELL,
        /** Defined name is a cell range */
        RANGE,
        /** Defined name is a formula */
        FORMULA,
        /** Defined name is a constant value */
        CONSTANT

    }

    private String name;
    private NameType type;
    private Object value;
    private Worksheet localSheet;
    private Worksheet targetWorksheet;
    private String textValue;
    private String comment;


    public String getName() {
        return name;
    }

    public String getComment() {
        return comment;
    }

    public NameType getType() {
        return type;
    }

    public Object getValue(){
        return value;
    }

    public String getTextValue(){
        return textValue;
    }

    /**
     * Compares this instance with another {@link DefinedName} for ordering. The order is
     * determined by {@link DefinedName#getName()} (ordinal), then by scope (workbook scope sorts before any
     * worksheet-scoped name; for worksheet-scoped names by {@link Worksheet#getSheetID()}), then
     * by {@link DefinedName#getTextValue()}, then by {@link  DefinedName#getComment()}.
     *
     * @param o Other defined name, or null. A null comparand sorts after this instance.
     * @return Negative, zero, or positive integer following the standard {@link Comparator{T}} contract.
     */
    @Override
    public int compareTo(DefinedName o) {
        if (o == null)
        {
            return 1;
        }
        int cmp = Comparator.nullsFirst(String::compareTo).compare(name, o.name);
        if (cmp != 0)
        {
            return cmp;
        }
        cmp = type.compareTo(o.type);
        if (cmp != 0)
        {
            return cmp;
        }
        cmp = compareScope(localSheet, o.localSheet);
        if (cmp != 0)
        {
            return cmp;
        }
        cmp = compareScope(targetWorksheet, o.targetWorksheet);
        if (cmp != 0)
        {
            return cmp;
        }
        cmp = Comparator.nullsFirst(String::compareTo).compare(textValue, o.textValue);
        if (cmp != 0)
        {
            return cmp;
        }
        return Comparator.nullsFirst(String::compareTo).compare(comment, o.comment);
    }


    /**
     * Compares two scope worksheets for ordering. Workbook scope (null) sorts before any worksheet
     * scope; two worksheet scopes are ordered by their {@link Worksheet#getSheetId()}.
     *
     * @param left Left scope, or null for workbook scope.
     * @param right Right scope, or null for workbook scope.
     * @return Negative, zero, or positive integer following the standard ordering contract.
     */
    private static int compareScope(Worksheet left, Worksheet right)
    {
        if (left == right)
        {
            return 0;
        }
        if (left == null)
        {
            return -1;
        }
        if (right == null)
        {
            return 1;
        }
        return Comparator.nullsFirst(Integer::compareTo).compare(left.getSheetID(), right.getSheetID());
    }

}
