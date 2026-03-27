//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface to represent a protection password, either for workbooks or worksheets.
 * Implementations define the password hashing algorithm used.
 *
 * @see ch.rabanti.nanoxlsx4j.interfaces.writer.IPasswordWriter
 * @see ch.rabanti.nanoxlsx4j.interfaces.reader.IPasswordReader
 */
public interface IPassword {

    /**
     * Gets the stored password hash.
     *
     * @return the password hash string, or {@code null} / empty if no password is set
     */
    String getPasswordHash();

    /**
     * Sets the password hash directly (bypassing plain-text hashing).
     *
     * @param hash the hash value to store
     */
    void setPasswordHash(String hash);

    /**
     * Hashes and stores the supplied plain-text password.
     *
     * @param plainText password in plain text
     */
    void setPassword(String plainText);

    /**
     * Clears any previously stored password.
     */
    void unsetPassword();

    /**
     * Returns the stored password as plain text.
     *
     * @return the plain-text password, or {@code null} if none is set
     */
    String getPassword();

    /**
     * Returns whether a password has been set.
     *
     * @return {@code true} if a password is currently stored
     */
    boolean passwordIsSet();

    /**
     * Copies the password state from another {@link IPassword} instance.
     *
     * @param source the source instance to copy from
     */
    void copyFrom(IPassword source);
}
