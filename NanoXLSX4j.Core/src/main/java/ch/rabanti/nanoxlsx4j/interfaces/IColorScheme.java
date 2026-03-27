//STATUS: CHECKED
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.interfaces;

/**
 * Interface to represent a color scheme that consists of 12 colors ({@link IColor}).
 *
 * @see IColor
 */
public interface IColorScheme {

    /** Gets the Dark 1 (dk1) color of the color scheme. */
    IColor getDark1();
    /** Sets the Dark 1 (dk1) color of the color scheme. */
    void setDark1(IColor color);

    /** Gets the Light 1 (lt1) color of the color scheme. */
    IColor getLight1();
    /** Sets the Light 1 (lt1) color of the color scheme. */
    void setLight1(IColor color);

    /** Gets the Dark 2 (dk2) color of the color scheme. */
    IColor getDark2();
    /** Sets the Dark 2 (dk2) color of the color scheme. */
    void setDark2(IColor color);

    /** Gets the Light 2 (lt2) color of the color scheme. */
    IColor getLight2();
    /** Sets the Light 2 (lt2) color of the color scheme. */
    void setLight2(IColor color);

    /** Gets the Accent 1 (accent1) color of the color scheme. */
    IColor getAccent1();
    /** Sets the Accent 1 (accent1) color of the color scheme. */
    void setAccent1(IColor color);

    /** Gets the Accent 2 (accent2) color of the color scheme. */
    IColor getAccent2();
    /** Sets the Accent 2 (accent2) color of the color scheme. */
    void setAccent2(IColor color);

    /** Gets the Accent 3 (accent3) color of the color scheme. */
    IColor getAccent3();
    /** Sets the Accent 3 (accent3) color of the color scheme. */
    void setAccent3(IColor color);

    /** Gets the Accent 4 (accent4) color of the color scheme. */
    IColor getAccent4();
    /** Sets the Accent 4 (accent4) color of the color scheme. */
    void setAccent4(IColor color);

    /** Gets the Accent 5 (accent5) color of the color scheme. */
    IColor getAccent5();
    /** Sets the Accent 5 (accent5) color of the color scheme. */
    void setAccent5(IColor color);

    /** Gets the Accent 6 (accent6) color of the color scheme. */
    IColor getAccent6();
    /** Sets the Accent 6 (accent6) color of the color scheme. */
    void setAccent6(IColor color);

    /** Gets the Hyperlink (hlink) color of the color scheme. */
    IColor getHyperlink();
    /** Sets the Hyperlink (hlink) color of the color scheme. */
    void setHyperlink(IColor color);

    /** Gets the Followed Hyperlink (folHlink) color of the color scheme. */
    IColor getFollowedHyperlink();
    /** Sets the Followed Hyperlink (folHlink) color of the color scheme. */
    void setFollowedHyperlink(IColor color);
}
