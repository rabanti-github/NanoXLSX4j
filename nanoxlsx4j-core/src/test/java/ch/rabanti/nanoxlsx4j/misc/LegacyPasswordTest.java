/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.misc;

import ch.rabanti.nanoxlsx4j.LegacyPassword;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.EnumSource;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class LegacyPasswordTest {

    @ParameterizedTest
    @DisplayName("Test of the generateLegacyPasswordHash function (legacy)")
    @CsvSource(value = {
            "x|CEBA",
            "Test@1-2,3!|F767",
            "' '|CE0A",
            "''|''",
            "NULL|''"
    }, delimiter = '|', nullValues = "NULL")
    public void generatePasswordHashTest(String givenPassword, String expectedHash) {
        String hash = LegacyPassword.generateLegacyPasswordHash(givenPassword);
        assertEquals(expectedHash, hash);
    }

    @ParameterizedTest
    @DisplayName("Test of the LegacyPassword constructor with arguments (legacy)")
    @EnumSource(LegacyPassword.PasswordType.class)
    public void constructorTest(LegacyPassword.PasswordType type) {
        LegacyPassword password = new LegacyPassword(type);
        assertNotNull(password);
        assertEquals(type, password.getType());
    }

    @ParameterizedTest
    @DisplayName("Test of the getType and setType methods (legacy)")
    @CsvSource({
            "WORKBOOK_PROTECTION, WORKSHEET_PROTECTION",
            "WORKSHEET_PROTECTION, WORKBOOK_PROTECTION"
    })
    public void passwordTypeTest(LegacyPassword.PasswordType initialType, LegacyPassword.PasswordType type) {
        LegacyPassword password = new LegacyPassword(initialType);
        assertEquals(initialType, password.getType());
        password.setType(type);
        assertEquals(type, password.getType());
    }

    @ParameterizedTest
    @DisplayName("Test of the getPasswordHash and setPasswordHash methods (legacy)")
    @CsvSource(value = {"CEBA", "''", "NULL", "0000"}, nullValues = "NULL")
    public void passwordHashTest(String passwordHash) {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        assertNull(password.getPasswordHash());
        password.setPasswordHash(passwordHash);
        assertEquals(passwordHash, password.getPasswordHash());
    }

    @ParameterizedTest
    @DisplayName("Test of the getPassword and setPassword functions (legacy)")
    @CsvSource(value = {
            "test|test",
            "0123|0123",
            "#@éü|#@éü",
            "' '|' '",
            "NULL|NULL",
            "''|NULL"
    }, delimiter = '|', nullValues = "NULL")
    public void getAndSetPasswordTest(String givenPassword, String expectedPassword) {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        assertNull(password.getPassword());
        password.setPassword(givenPassword);
        assertEquals(expectedPassword, password.getPassword());
    }

    @ParameterizedTest
    @DisplayName("Test of the unsetPassword function (legacy)")
    @CsvSource(value = {
            "CEBA|true",
            "''|false",
            "#@éü|true",
            "NULL|false",
            "0000|true"
    }, delimiter = '|', nullValues = "NULL")
    public void unsetPasswordTest(String plainText, boolean expectedPasswordSet) {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        assertNull(password.getPasswordHash());
        password.setPassword(plainText);
        if (expectedPasswordSet) {
            assertTrue(password.passwordIsSet());
            assertEquals(plainText, password.getPassword());
            assertNotNull(password.getPasswordHash());
        } else {
            assertFalse(password.passwordIsSet());
            assertNull(password.getPassword());
            assertNull(password.getPasswordHash());
        }

        password.unsetPassword();
        assertFalse(password.passwordIsSet());
        assertNull(password.getPassword());
        assertNull(password.getPasswordHash());
    }

    @ParameterizedTest
    @DisplayName("Test of the passwordIsSet function (legacy)")
    @CsvSource(value = {
            "CEBA|true",
            "''|false",
            "#@éü|true",
            "NULL|false",
            "0000|true"
    }, delimiter = '|', nullValues = "NULL")
    public void passwordIsSetTest(String passwordHash, boolean expectedPasswordSet) {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password.setPasswordHash(passwordHash);
        assertEquals(expectedPasswordSet, password.passwordIsSet());
    }

    @ParameterizedTest
    @DisplayName("Test of the copyFromTest function (legacy)")
    @CsvSource(value = {"CEBA", "''", "#@éü", "NULL", "0000"}, nullValues = "NULL")
    public void copyFromTest(String plainText) {
        LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        source.setPassword(plainText);
        LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        assertFalse(source.equals(target));
        target.copyFrom(source);
        assertTrue(source.equals(target));
    }

    @Test
    @DisplayName("Test of the hashCode function (legacy)")
    public void getHashCodeTest() {
        LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password1.setPassword("test");
        LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password2.setPassword("test");
        LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        password3.setPassword(null);
        assertEquals(password1.hashCode(), password2.hashCode());
        assertNotEquals(password1.hashCode(), password3.hashCode());
    }

    @Test
    @DisplayName("Test of the equals function (legacy)")
    public void equalsTest() {
        LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password1.setPassword("test");
        LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password2.setPassword("test");
        LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        password3.setPassword(null);
        assertTrue(password1.equals(password2));
        assertFalse(password1.equals(password3));
    }

    // Java-specific regression tests for owned-buffer cleanup; reference tests above remain unchanged.
    @Test
    @DisplayName("Test of clearing password buffers on replacement and unset (Java-specific)")
    public void replacementClearsOldBufferTest() throws Exception {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password.setPassword("old");
        char[] oldBuffer = passwordBuffer(password);
        password.setPassword("new");
        assertArrayEquals(new char[oldBuffer.length], oldBuffer);
        assertEquals("new", password.getPassword());
        char[] newBuffer = passwordBuffer(password);
        password.unsetPassword();
        assertArrayEquals(new char[newBuffer.length], newBuffer);
        assertNull(password.getPassword());
        password.unsetPassword();
        assertNull(password.getPasswordHash());
    }

    @Test
    @DisplayName("Test of clearing and independently copying password buffers (Java-specific)")
    public void copyClearsTargetAndOwnsBufferTest() throws Exception {
        LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
        source.setPassword("source");
        LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        target.setPassword("target");
        char[] oldBuffer = passwordBuffer(target);
        target.copyFrom(source);
        assertArrayEquals(new char[oldBuffer.length], oldBuffer);
        assertNotSame(passwordBuffer(source), passwordBuffer(target));
        source.unsetPassword();
        assertEquals("source", target.getPassword());
        assertEquals(LegacyPassword.PasswordType.WORKSHEET_PROTECTION, target.getType());
    }

    @Test
    @DisplayName("Test of preserving the password when copied from itself (Java-specific)")
    public void selfCopyPreservesPasswordTest() throws Exception {
        LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        password.setPassword("test");
        String hash = password.getPasswordHash();
        char[] oldBuffer = passwordBuffer(password);
        password.copyFrom(password);
        assertArrayEquals(new char[oldBuffer.length], oldBuffer);
        assertEquals("test", password.getPassword());
        assertEquals(hash, password.getPasswordHash());
    }

    @Test
    @DisplayName("Test of honoring an overridden getPassword method in copyFrom (Java-specific)")
    public void copyHonorsSubclassGetterTest() {
        LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION) {
            @Override
            public String getPassword() {
                return "overridden";
            }
        };
        source.setPassword("internal");
        LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        target.copyFrom(source);
        assertEquals("overridden", target.getPassword());
        assertEquals(source.getPasswordHash(), target.getPasswordHash());
    }

    @Test
    @DisplayName("Test of preserving the target state when copyFrom fails (Java-specific)")
    public void failedCopyPreservesTargetTest() {
        LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION) {
            @Override
            public String getPassword() {
                throw new IllegalStateException("Cannot read password");
            }
        };
        LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
        target.setPassword("retained");
        String hash = target.getPasswordHash();
        assertThrows(IllegalStateException.class, () -> target.copyFrom(source));
        assertEquals("retained", target.getPassword());
        assertEquals(hash, target.getPasswordHash());
        assertEquals(LegacyPassword.PasswordType.WORKBOOK_PROTECTION, target.getType());
    }

    private static char[] passwordBuffer(LegacyPassword password) throws Exception {
        // Inspect the original array to verify clearing, without exposing it in the production API.
        var field = LegacyPassword.class.getDeclaredField("password");
        field.setAccessible(true);
        return (char[]) field.get(password);
    }
}
