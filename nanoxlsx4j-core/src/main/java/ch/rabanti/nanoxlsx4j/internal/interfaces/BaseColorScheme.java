/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.internal.interfaces;

/**
 * Interface to represent a color scheme that consists of 12 colors ({@link BaseColor}).
 */
public interface BaseColorScheme {

    /**
     * Gets the Dark 1 (dk1) color of the color scheme.
     *
     * @return Dark 1 color
     */
    BaseColor getDark1();

    /**
     * Sets the Dark 1 (dk1) color of the color scheme.
     *
     * @param dark1 Dark 1 color
     */
    void setDark1(BaseColor dark1);

    /**
     * Gets the Light 1 (lt1) color of the color scheme.
     *
     * @return Light 1 color
     */
    BaseColor getLight1();

    /**
     * Sets the Light 1 (lt1) color of the color scheme.
     *
     * @param light1 Light 1 color
     */
    void setLight1(BaseColor light1);

    /**
     * Gets the Dark 2 (dk2) color of the color scheme.
     *
     * @return Dark 2 color
     */
    BaseColor getDark2();

    /**
     * Sets the Dark 2 (dk2) color of the color scheme.
     *
     * @param dark2 Dark 2 color
     */
    void setDark2(BaseColor dark2);

    /**
     * Gets the Light 2 (lt2) color of the color scheme.
     *
     * @return Light 2 color
     */
    BaseColor getLight2();

    /**
     * Sets the Light 2 (lt2) color of the color scheme.
     *
     * @param light2 Light 2 color
     */
    void setLight2(BaseColor light2);

    /**
     * Gets the Accent 1 (accent1) color of the color scheme.
     *
     * @return Accent 1 color
     */
    BaseColor getAccent1();

    /**
     * Sets the Accent 1 (accent1) color of the color scheme.
     *
     * @param accent1 Accent 1 color
     */
    void setAccent1(BaseColor accent1);

    /**
     * Gets the Accent 2 (accent2) color of the color scheme.
     *
     * @return Accent 2 color
     */
    BaseColor getAccent2();

    /**
     * Sets the Accent 2 (accent2) color of the color scheme.
     *
     * @param accent2 Accent 2 color
     */
    void setAccent2(BaseColor accent2);

    /**
     * Gets the Accent 3 (accent3) color of the color scheme.
     *
     * @return Accent 3 color
     */
    BaseColor getAccent3();

    /**
     * Sets the Accent 3 (accent3) color of the color scheme.
     *
     * @param accent3 Accent 3 color
     */
    void setAccent3(BaseColor accent3);

    /**
     * Gets the Accent 4 (accent4) color of the color scheme.
     *
     * @return Accent 4 color
     */
    BaseColor getAccent4();

    /**
     * Sets the Accent 4 (accent4) color of the color scheme.
     *
     * @param accent4 Accent 4 color
     */
    void setAccent4(BaseColor accent4);

    /**
     * Gets the Accent 5 (accent5) color of the color scheme.
     *
     * @return Accent 5 color
     */
    BaseColor getAccent5();

    /**
     * Sets the Accent 5 (accent5) color of the color scheme.
     *
     * @param accent5 Accent 5 color
     */
    void setAccent5(BaseColor accent5);

    /**
     * Gets the Accent 6 (accent6) color of the color scheme.
     *
     * @return Accent 6 color
     */
    BaseColor getAccent6();

    /**
     * Sets the Accent 6 (accent6) color of the color scheme.
     *
     * @param accent6 Accent 6 color
     */
    void setAccent6(BaseColor accent6);

    /**
     * Gets the Hyperlink (hlink) color of the color scheme.
     *
     * @return Hyperlink color
     */
    BaseColor getHyperlink();

    /**
     * Sets the Hyperlink (hlink) color of the color scheme.
     *
     * @param hyperlink Hyperlink color
     */
    void setHyperlink(BaseColor hyperlink);

    /**
     * Gets the Followed Hyperlink (folHlink) color of the color scheme.
     *
     * @return Followed Hyperlink color
     */
    BaseColor getFollowedHyperlink();

    /**
     * Sets the Followed Hyperlink (folHlink) color of the color scheme.
     *
     * @param followedHyperlink Followed Hyperlink color
     */
    void setFollowedHyperlink(BaseColor followedHyperlink);
}
