package ch.rabanti.nanoxlsx4j.cells.types;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.math.BigDecimal;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.Cell;

public class ConvertArrayTest {

	@DisplayName("Test of the convertArray method on boolean")
	@Test()
	void convertBooleanArrayTest() {
		Boolean[] array = new Boolean[] { true, true, false, true, false };
		assertArray(array, Boolean.class);
	}

	@DisplayName("Test of the convertArray method on byte")
	@Test()
	void convertByteArrayTest() {
		Byte[] array = new Byte[] { 12, 55, 127, -128, -1, 0, Byte.MAX_VALUE, Byte.MIN_VALUE };
		assertArray(array, Byte.class);
	}

	@DisplayName("Test of the convertArray method on BigDecimal")
	@Test()
	void convertBigDecimalArrayTest() {
		BigDecimal[] array = new BigDecimal[4];
		array[0] = BigDecimal.valueOf(0d);
		array[1] = BigDecimal.valueOf(2222L);
		array[2] = BigDecimal.valueOf(-0.0001);
		array[3] = BigDecimal.valueOf(1.01);
		assertArray(array, BigDecimal.class);
	}

	@DisplayName("Test of the convertArray method on double")
	@Test()
	void convertDoubleArrayTest() {
		Double[] array = new Double[] { 0d, 11.7d, 0.00001d, -22.5d, 100d, -99d, Double.MAX_VALUE, Double.MIN_VALUE };
		assertArray(array, Double.class);
	}

	@DisplayName("Test of the convertArray method on float")
	@Test()
	void convertFloatArrayTest() {
		Float[] array = new Float[] { 0f, 11.7f, 0.00001f, -22.5f, 100f, -99f, Float.MAX_VALUE, Float.MIN_VALUE };
		assertArray(array, Float.class);
	}

	@DisplayName("Test of the convertArray method on int")
	@Test()
	void convertIntArrayTest() {
		Integer[] array = new Integer[] { 12, 55, -1, 0, Integer.MAX_VALUE, Integer.MIN_VALUE };
		assertArray(array, Integer.class);
	}

	@DisplayName("Test of the convertArray method on long")
	@Test()
	void convertLongArrayTest() {
		Long[] array = new Long[] {
				12L,
				55L, -1L,
				0L, Long.MAX_VALUE, Long.MIN_VALUE };
		assertArray(array, Long.class);
	}

	@DisplayName("Test of the convertArray method on short")
	@Test()
	void convertShortArrayTest() {
		Short[] array = new Short[] { 12, 55, -1, 0, Short.MAX_VALUE, Short.MIN_VALUE };
		assertArray(array, Short.class);
	}

	@DisplayName("Test of the convertArray method on Date")
	@Test()
	void convertDateArrayTest() {
		Calendar calendarInstance = Calendar.getInstance();
		Date[] array = new Date[4];
		calendarInstance.set(1901, 01, 12, 12, 12, 12);
		array[0] = Date.from(calendarInstance.toInstant());
		calendarInstance.set(2200, 01, 12, 12, 12, 12);
		array[1] = Date.from(calendarInstance.toInstant());
		calendarInstance.set(2020, 11, 12);
		array[2] = Date.from(calendarInstance.toInstant());
		calendarInstance.set(1950, 5, 1);
		array[3] = Date.from(calendarInstance.toInstant());
		assertArray(array, Date.class);
	}

	@DisplayName("Test of the convertArray method on LocalTime")
	@Test()
	void convertLocalTimeArrayTest() {
		LocalTime[] array = new LocalTime[4];
		array[0] = LocalTime.of(0, 0, 0);
		array[1] = LocalTime.of(12, 10, 50);
		array[2] = LocalTime.of(23, 59, 59);
		array[3] = LocalTime.of(11, 11, 11);
		assertArray(array, LocalTime.class);
	}

	@DisplayName("Test of the convertArray method on nested Cell objects")
	@Test()
	void convertCellArrayTest() {
		Cell[] array = new Cell[4];
		array[0] = new Cell("", Cell.CellType.STRING);
		array[1] = new Cell("test", Cell.CellType.STRING);
		array[2] = new Cell("x", Cell.CellType.STRING);
		array[3] = new Cell(" ", Cell.CellType.STRING);
		assertArray(array,
				String.class,
				new String[] { "",
						"test",
						"x",
						" " });
	}

	@DisplayName("Test of the convertArray method on string")
	@Test()
	void convertStringArrayTest() {
		String[] array = new String[] { "",
				"test",
				"X",
				"Ø", null, " " };
		assertArray(array, String.class);
	}

	@DisplayName("Test of the convertArray method on other object types")
	@Test()
	void convertObjectArrayTest() {
		DummyArrayClass[] array = new DummyArrayClass[4];
		array[0] = new DummyArrayClass("");
		array[1] = new DummyArrayClass(null);
		array[2] = new DummyArrayClass(" ");
		array[3] = new DummyArrayClass("test");
		String[] actualValues = new String[array.length];
		for (int i = 0; i < actualValues.length; i++) {
			actualValues[i] = array[i].toString();
		}
		assertArray(array, String.class, actualValues);
	}

	@DisplayName("Test of the convertArray method on null and empty lists")
	@Test()
	void convertObjectArrayEmptyTest() {
		List<String> nullList = null;
		List<Cell> cells = Cell.convertArray(nullList);
		assertEquals(0, cells.size());

		List<String> emptyList = new ArrayList<>();
		List<Cell> cells2 = Cell.convertArray(emptyList);
		assertEquals(0, cells2.size());
	}

	private static <T> void assertArray(T[] array, Class<?> expectedValueType) {
		assertArray(array, expectedValueType, null);
	}

	private static <T> void assertArray(T[] array, Class<?> expectedValueType, Object[] actualValues) {
		List<T> list = new ArrayList<>();
		Collections.addAll(list, array);
		List<Cell> cells = Cell.convertArray(list);
		assertNotNull(cells);
		assertEquals(array.length, cells.size());
		for (int i = 0; i < array.length; i++) {
			Cell cell = cells.get(i);
			if (cell.getValue() != null) {
				assertEquals(expectedValueType, cell.getValue().getClass());
			}
			if (actualValues == null) {
				assertEquals(array[i], cell.getValue());
			}
			else {
				assertEquals(actualValues[i], cell.getValue());
			}

		}
	}

	public static class DummyArrayClass {
		private String value;

		public String getValue() {
			return value;
		}

		public void setValue(String value) {
			this.value = value;
		}

		public DummyArrayClass(String value) {
			this.value = value;
		}

		@Override
		public boolean equals(Object obj) {
			if (obj == null || !obj.getClass().equals(DummyArrayClass.class)) {
				return false;
			}
			return this.value.equals(((DummyArrayClass) obj).value);
		}

		@Override
		public String toString() {
			return this.value;
		}
	}

}
