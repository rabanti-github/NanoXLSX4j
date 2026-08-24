/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.themes;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColor;
import ch.rabanti.nanoxlsx4j.internal.interfaces.BaseColorScheme;

/**
 * Class representing a color scheme, used in a theme
 */
public class ColorScheme implements BaseColorScheme {

    private String name;
    private BaseColor dark1;
    private BaseColor dark2;
    private BaseColor light1;
    private BaseColor light2;
    private BaseColor accent1;
    private BaseColor accent2;
    private BaseColor accent3;
    private BaseColor accent4;
    private BaseColor accent5;
    private BaseColor accent6;
    private BaseColor hyperlink;
    private BaseColor followedHyperlink;

    /**
     * Gets the name of the color scheme
     *
     * @return Name of the scheme
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the name of the color scheme
     *
     * @param name Name of the scheme
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Gets the Theme color Dark 1 (dk1) attribute of a theme
     *
     * @return Dark 1 color
     */
    @Override
    public BaseColor getDark1() {
        return dark1;
    }

    /**
     * Sets the Dark 1 (dk1) attribute of a theme
     *
     * @param dark1 Dark 1 color
     */
    @Override
    public void setDark1(BaseColor dark1) {
        this.dark1 = dark1;
    }

    /**
     * Gets the Theme color Light 1 (lt1) attribute of a theme
     *
     * @return Light 1 color
     */
    @Override
    public BaseColor getLight1() {
        return light1;
    }

    /**
     * Sets the Light 1 (lt1) attribute of a theme
     *
     * @param light1 Light 1 color
     */
    @Override
    public void setLight1(BaseColor light1) {
        this.light1 = light1;
    }

    /**
     * Gets the Theme color Dark 2 (dk2) attribute of a theme
     *
     * @return Dark 2 color
     */
    @Override
    public BaseColor getDark2() {
        return dark2;
    }

    /**
     * Sets the Dark 2 (dk2) attribute of a theme
     *
     * @param dark2 Dark 2 color
     */
    @Override
    public void setDark2(BaseColor dark2) {
        this.dark2 = dark2;
    }

    /**
     * Gets the Theme color Light 2 (lt2) attribute of a theme
     *
     * @return Light 2 color
     */
    @Override
    public BaseColor getLight2() {
        return light2;
    }

    /**
     * Sets the Light 2 (lt2) attribute of a theme
     *
     * @param light2 Light 2 color
     */
    @Override
    public void setLight2(BaseColor light2) {
        this.light2 = light2;
    }

    /**
     * Gets the Theme color Accent 1 (accent1) attribute of a theme
     *
     * @return Accent 1 color
     */
    @Override
    public BaseColor getAccent1() {
        return accent1;
    }

    /**
     * Sets the Accent 1 (accent1) attribute of a theme
     *
     * @param accent1 Accent 1 color
     */
    @Override
    public void setAccent1(BaseColor accent1) {
        this.accent1 = accent1;
    }

    /**
     * Gets the Theme color Accent 2 (accent2) attribute of a theme
     *
     * @return Accent 2 color
     */
    @Override
    public BaseColor getAccent2() {
        return accent2;
    }

    /**
     * Sets the Accent 2 (accent2) attribute of a theme
     *
     * @param accent2 Accent 2 color
     */
    @Override
    public void setAccent2(BaseColor accent2) {
        this.accent2 = accent2;
    }

    /**
     * Gets the Theme color Accent 3 (accent3) attribute of a theme
     *
     * @return Accent 3 color
     */
    @Override
    public BaseColor getAccent3() {
        return accent3;
    }

    /**
     * Sets the Accent 3 (accent3) attribute of a theme
     *
     * @param accent3 Accent 3 color
     */
    @Override
    public void setAccent3(BaseColor accent3) {
        this.accent3 = accent3;
    }

    /**
     * Gets the Theme color Accent 4 (accent4) attribute of a theme
     *
     * @return Accent 4 color
     */
    @Override
    public BaseColor getAccent4() {
        return accent4;
    }

    /**
     * Sets the Accent 4 (accent4) attribute of a theme
     *
     * @param accent4 Accent 4 color
     */
    @Override
    public void setAccent4(BaseColor accent4) {
        this.accent4 = accent4;
    }

    /**
     * Gets the Theme color Accent 5 (accent5) attribute of a theme
     *
     * @return Accent 5 color
     */
    @Override
    public BaseColor getAccent5() {
        return accent5;
    }

    /**
     * Sets the Accent 5 (accent5) attribute of a theme
     *
     * @param accent5 Accent 5 color
     */
    @Override
    public void setAccent5(BaseColor accent5) {
        this.accent5 = accent5;
    }

    /**
     * Gets the Theme color Accent 6 (accent6) attribute of a theme
     *
     * @return Accent 6 color
     */
    @Override
    public BaseColor getAccent6() {
        return accent6;
    }

    /**
     * Sets the Accent 6 (accent6) attribute of a theme
     *
     * @param accent6 Accent 6 color
     */
    @Override
    public void setAccent6(BaseColor accent6) {
        this.accent6 = accent6;
    }

    /**
     * Gets the Theme color Hyperlink (hlink) attribute of a theme
     *
     * @return Hyperlink color
     */
    @Override
    public BaseColor getHyperlink() {
        return hyperlink;
    }

    /**
     * Sets the Hyperlink (hlink) attribute of a theme
     *
     * @param hyperlink Hyperlink color
     */
    @Override
    public void setHyperlink(BaseColor hyperlink) {
        this.hyperlink = hyperlink;
    }

    /**
     * Gets the Theme color Followed Hyperlink (folHlink) attribute of a theme
     *
     * @return Followed Hyperlink color
     */
    @Override
    public BaseColor getFollowedHyperlink() {
        return followedHyperlink;
    }

    /**
     * Sets the Followed Hyperlink (folHlink) attribute of a theme
     *
     * @param followedHyperlink Followed Hyperlink color
     */
    @Override
    public void setFollowedHyperlink(BaseColor followedHyperlink) {
        this.followedHyperlink = followedHyperlink;
    }

    /**
     * Default constructor
     *
     * <p>
     * Remarks: The constructor does not initialize any of the color properties. A workbook may become invalid on
     * saving, if any of the values are remaining null or undefined. This has to be maintained manually after
     * initialization
     * </p>
     */
    public ColorScheme() {
        // NoOp
    }

    /**
     * Returns whether two instances are the same
     *
     * @param o the reference object with which to compare.
     * @return True if this instance and the other are the same
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof ColorScheme that)) {
            return false;
        }

        return Objects.equals(name, that.name) && Objects.equals(dark1, that.dark1) &&
                Objects.equals(dark2, that.dark2) && Objects.equals(light1, that.light1) &&
                Objects.equals(light2, that.light2) && Objects.equals(accent1, that.accent1) &&
                Objects.equals(accent2, that.accent2) && Objects.equals(accent3, that.accent3) &&
                Objects.equals(accent4, that.accent4) && Objects.equals(accent5, that.accent5) &&
                Objects.equals(accent6, that.accent6) && Objects.equals(hyperlink, that.hyperlink) &&
                Objects.equals(followedHyperlink, that.followedHyperlink);
    }

    /**
     * Returns a hash code for this instance.
     *
     * @return A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash
     * table.
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(name);
        result = 31 * result + Objects.hashCode(dark1);
        result = 31 * result + Objects.hashCode(dark2);
        result = 31 * result + Objects.hashCode(light1);
        result = 31 * result + Objects.hashCode(light2);
        result = 31 * result + Objects.hashCode(accent1);
        result = 31 * result + Objects.hashCode(accent2);
        result = 31 * result + Objects.hashCode(accent3);
        result = 31 * result + Objects.hashCode(accent4);
        result = 31 * result + Objects.hashCode(accent5);
        result = 31 * result + Objects.hashCode(accent6);
        result = 31 * result + Objects.hashCode(hyperlink);
        result = 31 * result + Objects.hashCode(followedHyperlink);
        return result;
    }
}
