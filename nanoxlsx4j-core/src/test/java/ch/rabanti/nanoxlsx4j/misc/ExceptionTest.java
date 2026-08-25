/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.exceptions.FormatException;
import ch.rabanti.nanoxlsx4j.exceptions.IOException;
import ch.rabanti.nanoxlsx4j.exceptions.NotSupportedContentException;
import ch.rabanti.nanoxlsx4j.exceptions.PackageException;
import ch.rabanti.nanoxlsx4j.exceptions.RangeException;
import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.exceptions.WorksheetException;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;

public class ExceptionTest {

    // For code coverage
    @Test
    @DisplayName("Test of the FormatException (summary)")
    public void formatExceptionTest() throws Exception {
        FormatException exception = new FormatException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new FormatException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);

        IllegalArgumentException inner = new IllegalArgumentException("inner message");
        exception = new FormatException("test", inner);
        assertEquals("test", exception.getMessage());
        assertNotNull(exception.getCause());
        assertEquals(IllegalArgumentException.class, exception.getCause().getClass());
        assertEquals("inner message", exception.getCause().getMessage());
    }

    @Test
    @DisplayName("Test of the  IOExceptio (summary)")
    public void ioExceptionTest() throws Exception {
        IOException exception = new IOException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new IOException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);

        IllegalArgumentException inner = new IllegalArgumentException("inner message");
        exception = new IOException("test", inner);
        assertEquals("test", exception.getMessage());
        assertNotNull(exception.getCause());
        assertEquals(IllegalArgumentException.class, exception.getCause().getClass());
        assertEquals("inner message", exception.getCause().getMessage());
    }

    @Test
    @DisplayName("Test of the RangeException (summary)")
    public void rangeExceptionTest() throws Exception {
        RangeException exception = new RangeException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new RangeException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);
    }

    @Test
    @DisplayName("Test of the  StyleException (summary)")
    public void styleExceptionTest() throws Exception {
        StyleException exception = new StyleException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new StyleException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);

        IllegalArgumentException inner = new IllegalArgumentException("inner message");
        exception = new StyleException("test", inner);
        assertEquals("test", exception.getMessage());
        assertNotNull(exception.getCause());
        assertEquals(IllegalArgumentException.class, exception.getCause().getClass());
        assertEquals("inner message", exception.getCause().getMessage());
    }

    @Test
    @DisplayName("Test of the WorksheetException (summary)")
    public void worksheetExceptionTest() throws Exception {
        WorksheetException exception = new WorksheetException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new WorksheetException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);
    }

    @Test
    @DisplayName("Test of the NotSupportedContentException (summary)")
    public void notSupportedContentExceptionTest() throws Exception {
        NotSupportedContentException exception = new NotSupportedContentException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new NotSupportedContentException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);

        IllegalArgumentException inner = new IllegalArgumentException("inner message");
        exception = new NotSupportedContentException("test", inner);
        assertEquals("test", exception.getMessage());
        assertNotNull(exception.getCause());
        assertEquals(IllegalArgumentException.class, exception.getCause().getClass());
        assertEquals("inner message", exception.getCause().getMessage());
    }

    @Test
    @DisplayName("Test of the PackageException (summary)")
    public void packageExceptionTest() throws Exception {
        PackageException exception = new PackageException();
        assertNull(exception.getMessage()); // Java exceptions do not generate a default message.
        assertNull(exception.getCause());

        exception = new PackageException("test");
        assertEquals("test", exception.getMessage());
        assertNull(exception.getCause());

        assertExceptionSerialization(exception);

        IllegalArgumentException inner = new IllegalArgumentException("inner message");
        exception = new PackageException("test", inner);
        assertEquals("test", exception.getMessage());
        assertNotNull(exception.getCause());
        assertEquals(IllegalArgumentException.class, exception.getCause().getClass());
        assertEquals("inner message", exception.getCause().getMessage());
    }

    public static <TException extends Exception> void assertExceptionSerialization(TException originalException)
        throws java.io.IOException, ClassNotFoundException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ObjectOutputStream objectOutput = new ObjectOutputStream(output)) {
            objectOutput.writeObject(originalException);
        }

        TException deserializedException;
        try (ObjectInputStream objectInput = new ObjectInputStream(new ByteArrayInputStream(output.toByteArray()))) {
            @SuppressWarnings("unchecked")
            TException value = (TException) objectInput.readObject();
            deserializedException = value;
        }

        assertEquals(originalException.getMessage(), deserializedException.getMessage());
    }
}
