package ch.rabanti.nanoxlsx4j.cells.types;

import ch.rabanti.nanoxlsx4j.Cell;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class OtherTypesCellTest {

    CellTypeUtils utils;

    public OtherTypesCellTest() {
        utils = new CellTypeUtils(Object.class);
    }

    @DisplayName("Unknown value cell test: Test of the cell values, as well as proper modification")
    @Test()
    void unknownClassesCellTest() {
        DummyClass obj1 = new DummyClass(1);
        Cell actualCell = new Cell(obj1, Cell.CellType.DEFAULT, utils.getCellAddress());
        assertEquals(DummyClass.PREFIX + "1", actualCell.getValue().toString());
        assertEquals(DummyClass.class, actualCell.getValue().getClass());
        assertEquals(Cell.CellType.STRING, actualCell.getDataType());
        actualCell.setValue(new DummyClass2(2));
        assertEquals(DummyClass2.PREFIX + "2", actualCell.getValue().toString());
        assertEquals(DummyClass2.class, actualCell.getValue().getClass()); // should return the new class type
    }

    public static class DummyClass {
        public static final String PREFIX = "DummyValue = ";
        private int number;

        public int getNumber() {
            return number;
        }

        public void setNumber(int number) {
            this.number = number;
        }

        public DummyClass(int number) {
            this.number = number;
        }

        @Override
        public String toString() {
            return PREFIX + Integer.toString(this.number);
        }
    }

    public static class DummyClass2 {
        public static final String PREFIX = "DummyValue2 = ";
        private int number;

        public int getNumber() {
            return number;
        }

        public void setNumber(int number) {
            this.number = number;
        }

        public DummyClass2(int number) {
            this.number = number;
        }

        @Override
        public String toString() {
            return PREFIX + Integer.toString(this.number);
        }
    }

}
