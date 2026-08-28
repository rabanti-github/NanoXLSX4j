/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Interface to represent a protection password, either for workbooks or worksheets. The implementations will define the
 * password algorithms
 */
public interface Password {

    /**
     * Gets the password hash
     *
     * @return Hash value
     */
    String getPasswordHash();

    /**
     * Sets the password hash
     *
     * @param hash Hash value
     */
    void setPasswordHash(String hash);

    /**
     * Sets the plain text password
     *
     * @param plainText Password in plain text
     */
    void setPassword(String plainText);

    /**
     * Unsets a previously defined password
     */
    void unsetPassword();

    /**
     * Gets the password as plain text
     *
     * @return Password as plain text
     */
    String getPassword();

    /**
     * Gets whether a password was set or not
     *
     * @return True if the password was set, otherwise false
     */
    boolean passwordIsSet();

    /**
     * Method to copy a password instance from another one
     *
     * @param passwordInstance Source instance
     */
    void copyFrom(Password passwordInstance);
}
