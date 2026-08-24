/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and
 * native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

package ch.rabanti.nanoxlsx4j.colors;

import java.util.Objects;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.internal.interfaces.TypedColor;
import ch.rabanti.nanoxlsx4j.utils.Validators;

public class SystemColor implements TypedColor<SystemColor.Value> {

    /**
     * Enum defining the available system colors
     */
    public enum Value {
        /** 3D Dark System Color: Specifies a Dark shadow color for three-dimensional display elements */
        THREE_DIMENSIONAL_DARK_SHADOW,
        /** 3D Light System Color: Specifies a Light color for three-dimensional display elements */
        THREE_DIMENSIONAL_LIGHT,
        /** Active Border System Color: Specifies an Active Window Border Color */
        ACTIVE_BORDER,
        /**
         * Active Caption System Color: Specifies the active window title bar color. In particular the left side color
         * in the color gradient of an active window's title bar if the gradient effect is enabled
         */
        ACTIVE_CAPTION,
        /**
         * Application Workspace System Color: Specifies the Background color of multiple document interface (MDI)
         * applications
         */
        APP_WORKSPACE,
        /** Background System Color: Specifies the desktop background color */
        BACKGROUND,
        /**
         * Button Face System Color: Specifies the face color for three-dimensional display elements and for dialog box
         * backgrounds
         */
        BUTTON_FACE,
        /** Button Highlight System Color: Specifies the highlight color for three-dimensional display elements */
        BUTTON_HIGHLIGHT,
        /**
         * Button Shadow System Color: Specifies the shadow color for three-dimensional display elements (for edges
         * facing away from the light source)
         */
        BUTTON_SHADOW,
        /** Button Text System Color: Specifies the color of text on push buttons */
        BUTTON_TEXT,
        /** Caption Text System Color: Specifies the color of text in the caption, size box, and scroll bar arrow box */
        CAPTION_TEXT,
        /**
         * Gradient Active Caption System Color: Specifies the right side color in the color gradient of an active
         * window's title bar
         */
        GRADIENT_ACTIVE_CAPTION,
        /**
         * Gradient Inactive Caption System Color:  Specifies the right side color in the color gradient of an inactive
         * window's title bar
         */
        GRADIENT_INACTIVE_CAPTION,
        /**
         * Gray Text System Color: Specifies a grayed (disabled) text. This color is set to 0 if the current display
         * driver does not support a solid gray color
         */
        GRAY_TEXT,
        /** Highlight System Color: Specifies the color of Item(s) selected in a control */
        HIGHLIGHT,
        /** Highlight Text System Color: Specifies the text color of item(s) selected in a control */
        HIGHLIGHT_TEXT,
        /** Hot Light System Color: Specifies the color for a hyperlink or hot-tracked item */
        HOT_LIGHT,
        /** Inactive Border System Color: Specifies the color of the Inactive window border */
        INACTIVE_BORDER,
        /**
         * Inactive Caption System Color: Specifies the color of the Inactive window caption. Specifies the left side
         * color in the color gradient of an inactive window's title bar if the gradient effect is enabled
         */
        INACTIVE_CAPTION,
        /** Inactive Caption Text System Color: Specifies the color of text in an inactive caption */
        INACTIVE_CAPTION_TEXT,
        /** Info Back System Color: Specifies the background color for tool tip controls */
        INFO_BACKGROUND,
        /** Info Text System Color: Specifies the text color for tool tip controls */
        INFO_TEXT,
        /** Menu System Color: Specifies the menu background color */
        MENU,
        /** Menu Bar System Color: Specifies the background color for the menu bar when menus appear as flat menus */
        MENU_BAR,
        /**
         * Menu Highlight System Color: Specifies the color used to highlight menu items when the menu appears as a flat
         * menu
         */
        MENU_HIGHLIGHT,
        /** Menu Text System Color: Specifies the color of Text in menus */
        MENU_TEXT,
        /** Scroll Bar System Color: Specifies the scroll bar gray area color */
        SCROLL_BAR,
        /** Window System Color: Specifies window background color */
        WINDOW,
        /** Window Frame System Color: Specifies the window frame color */
        WINDOW_FRAME,
        /** Window Text System Color: Specifies the color of text in windows */
        WINDOW_TEXT,
    }

    private String lastColor = "000000";
    private Value colorValue =  Value.WINDOW_TEXT;

    /**
     * Gets the enum value of the system color
     *
     * @return Color enum values
     */
    @Override
    public Value getColorValue() {
        return colorValue;
    }

    /**
     * Sets the enum value of the system color
     *
     * @param colorValue Color enum values
     */
    @Override
    public void setColorValue(Value colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Gets the internal OOXML string value of the enum, defined by {@link SystemColor#setColorValue(Value)} ()}
     *
     * @return String of the system color
     */
    @Override
    public String getStringValue() {
        return mapValueToString(colorValue);
    }

    /**
     * Gets the color value that was last computed by the generating application
     *
     * @return Last color as RGB String
     */
    public String getLastColor() {
        return lastColor;
    }

    /**
     * Sets the color value that was last computed by the generating application
     *
     * @param lastColor Last color as RGB String
     */
    public void setLastColor(String lastColor) {
        Validators.validateColor(lastColor, false);
        this.lastColor = lastColor;
    }

    /**
     * Default constructor
     */
    public SystemColor() {
    }

    /**
     * Constructor with value as parameter
     *
     * @param colorValue Color value of the system color
     */
    public SystemColor(Value colorValue) {
        this();
        this.colorValue = colorValue;
    }

    /**
     * Constructor with all parameters
     *
     * @param colorValue Color value of the system color
     * @param lastColor  Last computed value
     */
    public SystemColor(Value colorValue, String lastColor) {
        this(colorValue);
        setLastColor(lastColor); // validate
    }

    /**
     * Maps the enum value of the system color to the OOXML value
     *
     * @param value Enum value
     * @return String value that can be placed in an XML document
     * @throws StyleException Thrown if an invalid value was passed
     */
    private static String mapValueToString(Value value) throws StyleException {
        return switch (value) {
            case Value.THREE_DIMENSIONAL_DARK_SHADOW -> "3dDkShadow";
            case Value.THREE_DIMENSIONAL_LIGHT -> "3dLight";
            case Value.ACTIVE_BORDER -> "activeBorder";
            case Value.ACTIVE_CAPTION -> "activeCaption";
            case Value.APP_WORKSPACE -> "appWorkspace";
            case Value.BACKGROUND -> "background";
            case Value.BUTTON_FACE -> "btnFace";
            case Value.BUTTON_HIGHLIGHT -> "btnHighlight";
            case Value.BUTTON_SHADOW -> "btnShadow";
            case Value.BUTTON_TEXT -> "btnText";
            case Value.CAPTION_TEXT -> "captionText";
            case Value.GRADIENT_ACTIVE_CAPTION -> "gradientActiveCaption";
            case Value.GRADIENT_INACTIVE_CAPTION -> "gradientInactiveCaption";
            case Value.GRAY_TEXT -> "grayText";
            case Value.HIGHLIGHT -> "highlight";
            case Value.HIGHLIGHT_TEXT -> "highlightText";
            case Value.HOT_LIGHT -> "hotLight";
            case Value.INACTIVE_BORDER -> "inactiveBorder";
            case Value.INACTIVE_CAPTION -> "inactiveCaption";
            case Value.INACTIVE_CAPTION_TEXT -> "inactiveCaptionText";
            case Value.INFO_BACKGROUND -> "infoBk";
            case Value.INFO_TEXT -> "infoText";
            case Value.MENU -> "menu";
            case Value.MENU_BAR -> "menuBar";
            case Value.MENU_HIGHLIGHT -> "menuHighlight";
            case Value.MENU_TEXT -> "menuText";
            case Value.SCROLL_BAR -> "scrollBar";
            case Value.WINDOW -> "window";
            case Value.WINDOW_FRAME -> "windowFrame";
            case Value.WINDOW_TEXT -> "windowText";
            default -> throw new StyleException(value + " is not a valid system color value");
        };
    }

    // TODO mapStringToValue was originally internal in C# -> Check accessibility

    /**
     * Maps a OOXML string value (from an XML document) to the corresponding enum value
     *
     * @param value OOXML string value
     * @return Enum value
     * @throws StyleException Thrown if an invalid string value was passed
     */
    private static Value mapStringToValue(String value) throws StyleException {
        return switch (value) {
            case "3dDkShadow" -> Value.THREE_DIMENSIONAL_DARK_SHADOW;
            case "3dLight" -> Value.THREE_DIMENSIONAL_LIGHT;
            case "activeBorder" -> Value.ACTIVE_BORDER;
            case "activeCaption" -> Value.ACTIVE_CAPTION;
            case "appWorkspace" -> Value.APP_WORKSPACE;
            case "background" -> Value.BACKGROUND;
            case "btnFace" -> Value.BUTTON_FACE;
            case "btnHighlight" -> Value.BUTTON_HIGHLIGHT;
            case "btnShadow" -> Value.BUTTON_SHADOW;
            case "btnText" -> Value.BUTTON_TEXT;
            case "captionText" -> Value.CAPTION_TEXT;
            case "gradientActiveCaption" -> Value.GRADIENT_ACTIVE_CAPTION;
            case "gradientInactiveCaption" -> Value.GRADIENT_INACTIVE_CAPTION;
            case "grayText" -> Value.GRAY_TEXT;
            case "highlight" -> Value.HIGHLIGHT;
            case "highlightText" -> Value.HIGHLIGHT_TEXT;
            case "hotLight" -> Value.HOT_LIGHT;
            case "inactiveBorder" -> Value.INACTIVE_BORDER;
            case "inactiveCaption" -> Value.INACTIVE_CAPTION;
            case "inactiveCaptionText" -> Value.INACTIVE_CAPTION_TEXT;
            case "infoBk" -> Value.INFO_BACKGROUND;
            case "infoText" -> Value.INFO_TEXT;
            case "menu" -> Value.MENU;
            case "menuBar" -> Value.MENU_BAR;
            case "menuHighlight" -> Value.MENU_HIGHLIGHT;
            case "menuText" -> Value.MENU_TEXT;
            case "scrollBar" -> Value.SCROLL_BAR;
            case "window" -> Value.WINDOW;
            case "windowFrame" -> Value.WINDOW_FRAME;
            case "windowText" -> Value.WINDOW_TEXT;
            default -> throw new StyleException(value + " is not a valid system color value");
        };
    }

    /**
     * Determines whether the specified object is equal to the current system color instance
     *
     * @param o the reference object with which to compare.
     * @return True if both objects are equal
     */
    @Override
    public final boolean equals(Object o) {
        if (!(o instanceof SystemColor that)) {
            return false;
        }

        return Objects.equals(lastColor, that.lastColor) && colorValue == that.colorValue;
    }

    /**
     * Gets the hash code of the system color instance
     *
     * @return hash code
     */
    @Override
    public int hashCode() {
        int result = Objects.hashCode(lastColor);
        result = 31 * result + Objects.hashCode(colorValue);
        return result;
    }
}
