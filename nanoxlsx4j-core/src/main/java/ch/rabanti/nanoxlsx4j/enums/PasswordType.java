/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.enums;

/** Target type of a password. */
public enum PasswordType {
    /** Password is used to protect a workbook. */
    WORKBOOK_PROTECTION,
    /** Password is used to protect a worksheet. */
    WORKSHEET_PROTECTION
}
