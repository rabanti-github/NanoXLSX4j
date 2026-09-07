/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;
import ch.rabanti.nanoxlsx4j.utils.Comparators;
import ch.rabanti.nanoxlsx4j.utils.ParserUtils;

import java.util.Arrays;
import java.util.Locale;
import java.util.Objects;

/**
 * Represents a legacy password based on Excel's proprietary hashing algorithm.
 *
 * <p>The retained character buffer is cleared on replacement and when unset. This is best-effort memory hygiene,
 * not encrypted storage or protection against memory inspection. Input and returned strings cannot be cleared
 * by this instance. Call {@link #unsetPassword()} when the password is no longer needed. Instances are not
 * thread-safe.</p>
 */
public class LegacyPassword implements Password {

    private char[] password;
    private PasswordType type;
    private String passwordHash;

    /**
     * Gets the password hash
     *
     * @return Hash value
     */
    @Override
    public String getPasswordHash() {
        return passwordHash;
    }

    /**
     * Sets the password hash
     *
     * @param hash Hash value
     */
    @Override
    public void setPasswordHash(String hash) {
        this.passwordHash = hash;
    }

    /**
     * Sets the plain text password
     *
     * @param plainText Password in plain text
     */
    @Override
    public void setPassword(String plainText) {
        if (ParserUtils.isNullOrEmpty(plainText)) {
            unsetPassword();
            return;
        }
        String replacementHash = generateLegacyPasswordHash(plainText);
        char[] replacement = plainText.toCharArray();
        clearPasswordBuffer();
        this.password = replacement;
        this.passwordHash = replacementHash;
    }

    /**
     * Unsets a previously defined password
     */
    @Override
    public void unsetPassword() {
        clearPasswordBuffer();
        passwordHash = null;
    }

    /**
     * Gets the password as plain text
     *
     * @return Password as plain text
     */
    @Override
    public String getPassword() {
        if (password != null && password.length > 0) {
            return new String(password);
        }
        return null;
    }

    /**
     * Gets whether a password was set or not
     *
     * @return True if the password was set, otherwise false
     */
    @Override
    public boolean passwordIsSet() {
        return !ParserUtils.isNullOrEmpty(passwordHash);
    }

    /**
     * Method to copy a password instance from another one
     *
     * @param passwordInstance Source instance
     */
    @Override
    public void copyFrom(Password passwordInstance) {
        Objects.requireNonNull(passwordInstance, "passwordInstance");
        String replacementHash = passwordInstance.getPasswordHash();
        char[] replacement;
        // Only bypass getPassword() for the concrete class: subclasses may override its public contract.
        if (passwordInstance.getClass() == LegacyPassword.class) {
            char[] source = ((LegacyPassword) passwordInstance).password;
            replacement = source == null ? null : source.clone();
        } else {
            String plainText = passwordInstance.getPassword();
            replacement = ParserUtils.isNullOrEmpty(plainText) ? null : plainText.toCharArray();
        }
        // Prepare first, including for self-copy, so failures leave this instance unchanged.
        clearPasswordBuffer();
        this.password = replacement;
        this.passwordHash = replacementHash;
        if (passwordInstance instanceof LegacyPassword legacyPassword) {
            this.type = legacyPassword.type;
        }
    }

    /**
     * Target type of the password
     */
    public enum PasswordType {
        /**
         * Password is used to protect a workbook
         */
        WORKBOOK_PROTECTION,
        /**
         * Password is used to protect a worksheet
         */
        WORKSHEET_PROTECTION
    }

    /**
     * Creates a legacy password for the specified target type.
     *
     * @param passwordType Password target type
     */
    public LegacyPassword(PasswordType passwordType) {
        this.type = passwordType;
        this.passwordHash = null;
    }

    /**
     * Gets the current target type of the password instance.
     *
     * @return Current password target type
     */
    public PasswordType getType() {
        return type;
    }

    /**
     * Sets the current target type of the password instance.
     *
     * @param type Password target type
     */
    public void setType(PasswordType type) {
        this.type = type;
    }

    /**
     * Generates a legacy Excel password hash for workbook or worksheet protection.
     *
     * @param password Password string to hash
     * @return 16-bit hash as an uppercase hexadecimal string, or an empty string for a null or empty password
     */
    public static String generateLegacyPasswordHash(String password) {
        if (ParserUtils.isNullOrEmpty(password)) {
            return "";
        }
        int passwordLength = password.length();
        int passwordHash = 0;
        for (int i = passwordLength; i > 0; i--) {
            char character = password.charAt(i - 1);
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= character;
        }
        passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
        passwordHash ^= (0x8000 | ('N' << 8) | 'K');
        passwordHash ^= passwordLength;
        return Integer.toHexString(passwordHash).toUpperCase(Locale.ROOT);
    }

    private void clearPasswordBuffer() {
        if (password != null) {
            Arrays.fill(password, '\0');
            password = null;
        }
    }

    @Override
    public boolean equals(Object obj) {
        return obj instanceof LegacyPassword legacyPassword
                && Comparators.compareSecureStrings(this.password, legacyPassword.password)
                && type == legacyPassword.type
                && Objects.equals(passwordHash, legacyPassword.passwordHash);
    }

    @Override
    public int hashCode() {
        // The actual password is not considered since its hash is sufficient.
        return Objects.hash(type, passwordHash, passwordIsSet());
    }

}
