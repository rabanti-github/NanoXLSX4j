package ch.rabanti.nanoxlsx4j.misc;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;

public class ExceptionTest {
	// For code coverage
	@DisplayName("Test of the FormatException (summary)")
	@Test()
	void formatExceptionTest() {
		FormatException exception = new FormatException();
		assertNotEquals("", exception.getMessage()); // Gets a generated message my the base class
		assertNull(exception.getInnerException());

		exception = new FormatException("test");
		assertEquals("test", exception.getMessage());
		assertNull(exception.getInnerException());

		IllegalArgumentException inner = new IllegalArgumentException("inner message");
		exception = new FormatException("test", inner);
		assertEquals("test", exception.getMessage());
		assertNotNull(exception.getInnerException());
		assertEquals(IllegalArgumentException.class, exception.getInnerException().getClass());
		assertEquals("inner message", exception.getInnerException().getMessage());
	}

	@DisplayName("Test of the  IOExceptio (summary)")
	@Test()
	void iOExceptionTest() {
		IOException exception = new IOException();
		assertNotEquals("", exception.getMessage()); // Gets a generated message my the base class
		assertNull(exception.getInnerException());

		exception = new IOException("test");
		assertEquals("test", exception.getMessage());
		assertNull(exception.getInnerException());

		IllegalArgumentException inner = new IllegalArgumentException("inner message");
		exception = new IOException("test", inner);
		assertEquals("test", exception.getMessage());
		assertNotNull(exception.getInnerException());
		assertEquals(IllegalArgumentException.class, exception.getInnerException().getClass());
		assertEquals("inner message", exception.getInnerException().getMessage());
	}

	@DisplayName("Test of the RangeException (summary)")
	@Test()
	void rangeExceptionTest() {
		RangeException exception = new RangeException();
		assertNotEquals("", exception.getMessage()); // Gets a generated message my the base class

		exception = new RangeException("test");
		assertEquals("test", exception.getMessage());
	}

	@DisplayName("Test of the  StyleException (summary)")
	@Test()
	void styleExceptionTest() {
		StyleException exception = new StyleException();
		assertNotEquals("", exception.getMessage()); // Gets a generated message my the base class

		exception = new StyleException("test");
		assertEquals("test", exception.getMessage());

		IllegalArgumentException inner = new IllegalArgumentException("inner message");
		exception = new StyleException("test", inner);
		assertEquals("test", exception.getMessage());
	}

	@DisplayName("Test of the WorksheetException (summary)")
	@Test()
	void worksheetExceptionTest() {
		WorksheetException exception = new WorksheetException();
		assertNotEquals("", exception.getMessage()); // Gets a generated message my the base class

		exception = new WorksheetException("test");
		assertEquals("test", exception.getMessage());
	}

}
