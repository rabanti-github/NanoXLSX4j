//STATUS: CHECKED
package ch.rabanti.nanoxlsx4j.enums;

/**
 * Enum representing the target type of password protection.
 * <p>
 * Java equivalent of {@code NanoXLSX.Enums.Password.PasswordType} in the .NET reference.
 * </p>
 *
 * @see ch.rabanti.nanoxlsx4j.interfaces.IPassword
 * @see ch.rabanti.nanoxlsx4j.interfaces.writer.IPasswordWriter
 * @see ch.rabanti.nanoxlsx4j.interfaces.reader.IPasswordReader
 */
public enum PasswordType {

    /** Password is used to protect a workbook. */
    WORKBOOK_PROTECTION,

    /** Password is used to protect a worksheet. */
    WORKSHEET_PROTECTION
}
