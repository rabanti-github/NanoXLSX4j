/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.utils;

import ch.rabanti.nanoxlsx4j.enums.PasswordType;
import ch.rabanti.nanoxlsx4j.interfaces.IPassword;

/**
 * Implementation of {@link IPassword} using the legacy Excel password hashing algorithm.
 * <p>
 * Java equivalent of {@code NanoXLSX.Utils.LegacyPassword} in the .NET reference.
 * Password hashing is performed externally by the caller (Workbook) via {@link #setPasswordHash(String)}
 * until the hashing utility is moved to Core during v3 source migration.
 * </p>
 *
 * @see IPassword
 * @see PasswordType
 */
public class LegacyPassword implements IPassword {

    private final PasswordType type;
    private String plainText;
    private String passwordHash;

    /**
     * Constructor with password type
     *
     * @param type Target type of the password (workbook or worksheet protection)
     */
    public LegacyPassword(PasswordType type) {
        this.type = type;
    }

    /**
     * Gets the target type of this password
     *
     * @return Password type
     */
    public PasswordType getType() {
        return type;
    }

    /**
     * {@inheritDoc}
     * <p>
     * Note: password hashing is not performed here in the stub. Use
     * {@link #setPasswordHash(String)} to supply the computed hash.
     * </p>
     */
    @Override
    public void setPassword(String plainText) {
        this.plainText = plainText;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void setPasswordHash(String hash) {
        this.passwordHash = hash;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getPasswordHash() {
        return passwordHash;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public String getPassword() {
        return plainText;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean passwordIsSet() {
        return plainText != null && !plainText.isEmpty();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void unsetPassword() {
        this.plainText = null;
        this.passwordHash = null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void copyFrom(IPassword source) {
        this.plainText = source.getPassword();
        this.passwordHash = source.getPasswordHash();
    }
}
