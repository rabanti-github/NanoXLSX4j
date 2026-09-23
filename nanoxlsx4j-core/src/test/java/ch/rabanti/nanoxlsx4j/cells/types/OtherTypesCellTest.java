package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import ch.rabanti.nanoxlsx4j.Cell;

class OtherTypesCellTest {
    private final CellTypeUtils utils = new CellTypeUtils();

    @Test
    @DisplayName("Unknown value cell test: Test of the cell values, as well as proper modification")
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

    private static class DummyClass {
        static final String PREFIX = "DummyValue = ";
        private final int number;

        DummyClass(int number) {
            this.number = number;
        }

        @Override
        public String toString() {
            return PREFIX + number;
        }
    }

    private static class DummyClass2 {
        static final String PREFIX = "DummyValue2 = ";
        private final int number;

        DummyClass2(int number) {
            this.number = number;
        }

        @Override
        public String toString() {
            return PREFIX + number;
        }
    }
}
