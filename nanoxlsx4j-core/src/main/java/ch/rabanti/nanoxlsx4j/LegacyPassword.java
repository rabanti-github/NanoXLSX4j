package ch.rabanti.nanoxlsx4j;

import ch.rabanti.nanoxlsx4j.internal.interfaces.Password;

public class LegacyPassword implements Password {

    /**
     * Gets the password hash
     *
     * @return Hash value
     */
    @Override
    public String getPasswordHash() {
        return "";
    }

    /**
     * Sets the password hash
     *
     * @param hash Hash value
     */
    @Override
    public void setPasswordHash(String hash) {

    }

    /**
     * Sets the plain text password
     *
     * @param plainText Password in plain text
     */
    @Override
    public void setPassword(String plainText) {

    }

    /**
     * Unsets a previously defined password
     */
    @Override
    public void unsetPassword() {

    }

    /**
     * Gets the password as plain text
     *
     * @return Password as plain text
     */
    @Override
    public String getPassword() {
        return "";
    }

    /**
     * Gets whether a password was set or not
     *
     * @return True if the password was set, otherwise false
     */
    @Override
    public boolean passwordIsSet() {
        return false;
    }

    /**
     * Method to copy a password instance from another one
     *
     * @param passwordInstance Source instance
     */
    @Override
    public void copyFrom(Password passwordInstance) {

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

    public LegacyPassword(PasswordType passwordType) {
        // TODO implement
    }

}
