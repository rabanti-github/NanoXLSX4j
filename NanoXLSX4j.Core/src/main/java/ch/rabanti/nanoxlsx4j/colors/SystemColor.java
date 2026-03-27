//STATUS: REVISE ENUM
/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli @ 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.colors;

import ch.rabanti.nanoxlsx4j.exceptions.StyleException;
import ch.rabanti.nanoxlsx4j.interfaces.ITypedColor;

import java.util.Objects;

/**
 * Class representing a predefined system color for certain purposes or target areas in the UI.
 * <p>
 * Java equivalent of {@code NanoXLSX.Colors.SystemColor} in the .NET reference.
 * </p>
 */
public class SystemColor implements ITypedColor<SystemColor.Value> {

    /**
     * Enum defining the available system colors
     */
    public enum Value {
        /** 3D Dark System Color: Dark shadow for three-dimensional display elements */
        ThreeDimensionalDarkShadow,
        /** 3D Light System Color: Light color for three-dimensional display elements */
        ThreeDimensionalLight,
        /** Active Border System Color: Active Window Border Color */
        ActiveBorder,
        /** Active Caption System Color: Active window title bar color */
        ActiveCaption,
        /** Application Workspace System Color: Background of MDI applications */
        AppWorkspace,
        /** Background System Color: Desktop background color */
        Background,
        /** Button Face System Color: Face color for 3D elements and dialog backgrounds */
        ButtonFace,
        /** Button Highlight System Color: Highlight color for 3D display elements */
        ButtonHighlight,
        /** Button Shadow System Color: Shadow color for 3D display elements */
        ButtonShadow,
        /** Button Text System Color: Color of text on push buttons */
        ButtonText,
        /** Caption Text System Color: Color of text in caption, size box, and scroll bar arrow box */
        CaptionText,
        /** Gradient Active Caption System Color: Right side color in active window title bar gradient */
        GradientActiveCaption,
        /** Gradient Inactive Caption System Color: Right side color in inactive window title bar gradient */
        GradientInactiveCaption,
        /** Gray Text System Color: Grayed (disabled) text */
        GrayText,
        /** Highlight System Color: Color of selected items in a control */
        Highlight,
        /** Highlight Text System Color: Text color of selected items in a control */
        HighlightText,
        /** Hot Light System Color: Color for a hyperlink or hot-tracked item */
        HotLight,
        /** Inactive Border System Color: Color of the inactive window border */
        InactiveBorder,
        /** Inactive Caption System Color: Color of the inactive window caption */
        InactiveCaption,
        /** Inactive Caption Text System Color: Color of text in an inactive caption */
        InactiveCaptionText,
        /** Info Background System Color: Background color for tool tip controls */
        InfoBackground,
        /** Info Text System Color: Text color for tool tip controls */
        InfoText,
        /** Menu System Color: Menu background color */
        Menu,
        /** Menu Bar System Color: Background color for the menu bar in flat menu mode */
        MenuBar,
        /** Menu Highlight System Color: Color used to highlight menu items in flat menu mode */
        MenuHighlight,
        /** Menu Text System Color: Color of text in menus */
        MenuText,
        /** Scroll Bar System Color: Scroll bar gray area color */
        ScrollBar,
        /** Window System Color: Window background color */
        Window,
        /** Window Frame System Color: Window frame color */
        WindowFrame,
        /** Window Text System Color: Color of text in windows */
        WindowText
    }

    private Value colorValue = Value.WindowText;
    private String lastColor = "000000";

    /**
     * Gets the system color enum value
     *
     * @return System color enum value
     */
    @Override
    public Value getColorValue() {
        return colorValue;
    }

    /**
     * Sets the system color enum value
     *
     * @param value System color enum value
     */
    @Override
    public void setColorValue(Value value) {
        this.colorValue = value;
    }

    /**
     * Gets the OOXML string representation of the system color
     *
     * @return OOXML string value
     */
    @Override
    public String getStringValue() {
        return mapValueToString(colorValue);
    }

    /**
     * Gets the last computed ARGB hex value for this system color
     *
     * @return Hex color string (6 characters, without alpha)
     */
    public String getLastColor() {
        return lastColor;
    }

    /**
     * Sets the last computed color value
     *
     * @param lastColor Hex color string (6 characters, without alpha)
     * @throws StyleException Thrown if the value is not a valid 6-character hex string
     */
    public void setLastColor(String lastColor) {
        if (lastColor == null || lastColor.length() != 6 || !lastColor.matches("[0-9A-Fa-f]+")) {
            throw new StyleException("The lastColor value '" + lastColor + "' is invalid. It must be a 6-character hex string");
        }
        this.lastColor = lastColor;
    }

    /**
     * Default constructor
     */
    public SystemColor() {
    }

    /**
     * Constructor with a system color value
     *
     * @param value System color enum value
     */
    public SystemColor(Value value) {
        this.colorValue = value;
    }

    /**
     * Constructor with a system color value and last computed color
     *
     * @param value     System color enum value
     * @param lastColor Last computed ARGB hex color value (6 characters, without alpha)
     * @throws StyleException Thrown if the lastColor is not a valid 6-character hex string
     */
    public SystemColor(Value value, String lastColor) {
        this.colorValue = value;
        setLastColor(lastColor);
    }

    /**
     * Maps the system color enum value to its OOXML XML attribute string
     *
     * @param value Enum value
     * @return OOXML string representation
     * @throws StyleException Thrown if the enum value has no OOXML mapping
     */
    private static String mapValueToString(Value value) {
        switch (value) {
            case ThreeDimensionalDarkShadow:
                return "3dDkShadow";
            case ThreeDimensionalLight:
                return "3dLight";
            case ActiveBorder:
                return "activeBorder";
            case ActiveCaption:
                return "activeCaption";
            case AppWorkspace:
                return "appWorkspace";
            case Background:
                return "background";
            case ButtonFace:
                return "btnFace";
            case ButtonHighlight:
                return "btnHighlight";
            case ButtonShadow:
                return "btnShadow";
            case ButtonText:
                return "btnText";
            case CaptionText:
                return "captionText";
            case GradientActiveCaption:
                return "gradientActiveCaption";
            case GradientInactiveCaption:
                return "gradientInactiveCaption";
            case GrayText:
                return "grayText";
            case Highlight:
                return "highlight";
            case HighlightText:
                return "highlightText";
            case HotLight:
                return "hotLight";
            case InactiveBorder:
                return "inactiveBorder";
            case InactiveCaption:
                return "inactiveCaption";
            case InactiveCaptionText:
                return "inactiveCaptionText";
            case InfoBackground:
                return "infoBk";
            case InfoText:
                return "infoText";
            case Menu:
                return "menu";
            case MenuBar:
                return "menuBar";
            case MenuHighlight:
                return "menuHighlight";
            case MenuText:
                return "menuText";
            case ScrollBar:
                return "scrollBar";
            case Window:
                return "window";
            case WindowFrame:
                return "windowFrame";
            case WindowText:
                return "windowText";
            default:
                throw new StyleException(value + " is not a valid system color value");
        }
    }

    /**
     * Maps an OOXML string value (from an XML document) to the corresponding system color enum value.
     * <p>
     * Java equivalent of {@code internal static Value MapStringToValue(string)} in the .NET reference.
     * </p>
     *
     * @param value OOXML string value
     * @return System color enum value
     * @throws StyleException Thrown if the string has no enum mapping
     */
    public static Value mapStringToValue(String value) {
        switch (value) {
            case "3dDkShadow":
                return Value.ThreeDimensionalDarkShadow;
            case "3dLight":
                return Value.ThreeDimensionalLight;
            case "activeBorder":
                return Value.ActiveBorder;
            case "activeCaption":
                return Value.ActiveCaption;
            case "appWorkspace":
                return Value.AppWorkspace;
            case "background":
                return Value.Background;
            case "btnFace":
                return Value.ButtonFace;
            case "btnHighlight":
                return Value.ButtonHighlight;
            case "btnShadow":
                return Value.ButtonShadow;
            case "btnText":
                return Value.ButtonText;
            case "captionText":
                return Value.CaptionText;
            case "gradientActiveCaption":
                return Value.GradientActiveCaption;
            case "gradientInactiveCaption":
                return Value.GradientInactiveCaption;
            case "grayText":
                return Value.GrayText;
            case "highlight":
                return Value.Highlight;
            case "highlightText":
                return Value.HighlightText;
            case "hotLight":
                return Value.HotLight;
            case "inactiveBorder":
                return Value.InactiveBorder;
            case "inactiveCaption":
                return Value.InactiveCaption;
            case "inactiveCaptionText":
                return Value.InactiveCaptionText;
            case "infoBk":
                return Value.InfoBackground;
            case "infoText":
                return Value.InfoText;
            case "menu":
                return Value.Menu;
            case "menuBar":
                return Value.MenuBar;
            case "menuHighlight":
                return Value.MenuHighlight;
            case "menuText":
                return Value.MenuText;
            case "scrollBar":
                return Value.ScrollBar;
            case "window":
                return Value.Window;
            case "windowFrame":
                return Value.WindowFrame;
            case "windowText":
                return Value.WindowText;
            default:
                throw new StyleException(value + " is not a valid system color value");
        }
    }

    /**
     * Determines whether the specified object is equal to the current system color instance
     *
     * @param obj Other object to compare
     * @return True if both objects are equal
     */
    @Override
    public boolean equals(Object obj) {
        if (this == obj) return true;
        if (!(obj instanceof SystemColor)) return false;
        SystemColor other = (SystemColor) obj;
        return colorValue == other.colorValue && Objects.equals(lastColor, other.lastColor);
    }

    /**
     * Gets the hash code of the system color instance
     *
     * @return Hash code
     */
    @Override
    public int hashCode() {
        int hashCode = 1425985453;
        hashCode = hashCode * -1521134295 + colorValue.hashCode();
        hashCode = hashCode * -1521134295 + Objects.hashCode(lastColor);
        return hashCode;
    }
}
