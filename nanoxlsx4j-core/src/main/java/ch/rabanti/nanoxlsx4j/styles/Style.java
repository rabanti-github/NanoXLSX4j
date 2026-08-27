package ch.rabanti.nanoxlsx4j.styles;

// TODO implement
public class Style {

    @AppendAnnotation(nestedProperty = true)
    private CellXf currentCellXf;

    public CellXf getCurrentCellXf() {
        return currentCellXf;
    }

    public Style copyStyle(){
        return null;
    }
}
