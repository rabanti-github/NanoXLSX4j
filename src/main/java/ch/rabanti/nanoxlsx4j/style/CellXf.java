/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.style;

import ch.rabanti.nanoxlsx4j.exception.RangeException;

import static ch.rabanti.nanoxlsx4j.style.StyleManager.CELLXFPREFIX;

/**
 * Class representing an XF entry. The XF entry is used to make reference to other style instances like Border or Fill and for the positioning of the cell content
 * @author Raphael Stoeckli
 */
public class CellXf  extends AbstractStyle
{
// ### E N U M S ###
    /**
     * Enum for the horizontal alignment of a cell 
     */
     public enum HorizontalAlignValue
     {
         /**
         * Content will be aligned left
         */
         left(1),
         /**
         * Content will be aligned in the center
         */
         center(2),
         /**
         * Content will be aligned right
         */
         right(3),
         /**
         * Content will fill up the cell
         */
         fill(4),
         /**
         * justify alignment
         */
         justify(5),
         /**
         * General alignment
         */
         general(6),
         /**
         * Center continuous alignment
         */
         centerContinuous(7),
         /**
         * Distributed alignment
         */
         distributed(8),
         /**
         * No alignment. The alignment will not be used in a style
         */
         none(0);
         
        private final int value;
        HorizontalAlignValue(int value) { this.value = value; }
        public int getValue() { return value; }
         
     }

     /**
      * Enum for the vertical alignment of a cell 
      */
     public enum VerticalAlignValue
     {
         /**
         * Content will be aligned on the bottom (default)
         */
         bottom(1),
         /**
         * Content will be aligned on the top
         */
         top(2),
         /**
         * Content will be aligned in the center
         */
         center(3),
         /**
         * justify alignment
         */
         justify(4),
         /**
         * Distributed alignment
         */
         distributed(5),
         /**
         * No alignment. The alignment will not be used in a style
         */
         none(0);
         
        private final int value;
        VerticalAlignValue(int value) { this.value = value; }
        public int getValue() { return value; }
     }     

     /**
      * Enum for text break options
      */
     public enum TextBreakValue
     {
         /**
          * Word wrap is active
          */
         wrapText(1),
         /**
          * Text will be resized to fit the cell
          */
         shrinkToFit(2),
         /**
          * Text will overflow in cell
          */
         none(0);
        private final int value;
        TextBreakValue(int value) { this.value = value; }
        public int getValue() { return value; }
     }

     /**
      * Enum for the general text alignment direction
      */
     public enum TextDirectionValue
     {
         /**
          * Text direction is horizontal (default)
          */
         horizontal(0),
         /**
          * Text direction is vertical
          */
         vertical(1);
         
        private final int value;
        TextDirectionValue(int value) { this.value = value; }
        public int getValue() { return value; }
     }

// ### P R I V A T E  F I E L D S ###
     private int textRotation;
     private TextDirectionValue textDirection;           
     private HorizontalAlignValue horizontalAlign;
     private VerticalAlignValue verticalAlign;
     private TextBreakValue alignment;
     private boolean locked;
     private boolean hidden;
     private boolean forceApplyAlignment;

// ### G E T T E R S  &  S E T T E R S ###
     /**
      * Gets the text rotation in degrees (from +90 to -90)
      * @return Text rotation in degrees (from +90 to -90)
      */
     public int getTextRotation() {
         return textRotation;
     }

     /**
      * Sets the text rotation in degrees (from +90 to -90)
      * @param textRotation Text rotation in degrees (from +90 to -90)
      * @throws RangeException Thrown if the rotation angle is out of range
      */
     public void setTextRotation(int textRotation) {
         this.textRotation = textRotation;
         this.textDirection = TextDirectionValue.horizontal;
         calculateInternalRotation();
     }

     /**
      * Gets the direction of the text within the cell
      * @return Direction of the text within the cell
      */
     public TextDirectionValue getTextDirection() {
         return textDirection;
     }

     /**
      * Sets the direction of the text within the cell
      * @param textDirection Direction of the text within the cell
      * @throws RangeException Thrown if the text rotation and direction causes a conflict
      */
     public void setTextDirection(TextDirectionValue textDirection) {
         this.textDirection = textDirection;
         calculateInternalRotation();            
     }

     /**
      * Gets the horizontal alignment of the style
      * @return Horizontal alignment of the style
      */
     public HorizontalAlignValue getHorizontalAlign() {
         return horizontalAlign;
     }

     /**
      * Sets the horizontal alignment of the style
      * @param horizontalAlign Horizontal alignment of the style
      */
     public void setHorizontalAlign(HorizontalAlignValue horizontalAlign) {
         this.horizontalAlign = horizontalAlign;
     }

     /**
      * Gets the vertical alignment of the style
      * @return Vertical alignment of the style
      */
     public VerticalAlignValue getVerticalAlign() {
         return verticalAlign;
     }

     /**
      * Sets the vertical alignment of the style
      * @param verticalAlign Vertical alignment of the style
      */
     public void setVerticalAlign(VerticalAlignValue verticalAlign) {
         this.verticalAlign = verticalAlign;
     }

     /**
      * Gets the text break options of the style
      * @return Text break options of the style
      */
     public TextBreakValue getAlignment() {
         return alignment;
     }

     /**
      * Sets the text break options of the style
      * @param alignment Text break options of the style
      */
     public void setAlignment(TextBreakValue alignment) {
         this.alignment = alignment;
     }

    /**
     * Gets whether the locked property (used for locking / protection of cells or worksheets) will be defined in the XF entry of the style. If true, locked will be defined
     * @return If true, the style is used for locking / protection of cells or worksheets
     */
    public boolean isLocked() {
        return locked;
    }

    /**
     * Sets whether the locked property (used for locking / protection of cells or worksheets) will be defined in the XF entry of the style. If true, locked will be defined
     * @param locked If true, the style is used for locking / protection of cells or worksheets
     */
    public void setLocked(boolean locked) {
        this.locked = locked;
    }

    /**
     * Gets whether the hidden property (used for protection or hiding of cells) will be defined in the XF entry of the style. If true, hidden will be defined
     * @return If true, the style is used for hiding cell values / protection of cells
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the hidden property (used for protection or hiding of cells) will be defined in the XF entry of the style. If true, hidden will be defined
     * @param hidden If true, the style is used for hiding cell values / protection of cells
     */
    public void setHidden(boolean hidden) {
        this.hidden = hidden;
    }

    /**
     * Gets whether the applyAlignment property (used to merge cells) will be defined in the XF entry of the style. If true, applyAlignment will be defined
     * @return If true, the applyAlignment value of the style will be set to true (used to merge cells)
     */
    public boolean isForceApplyAlignment() {
        return forceApplyAlignment;
    }

    /**
     * Sets whether the applyAlignment property (used to merge cells) will be defined in the XF entry of the style. If true, applyAlignment will be defined
     * @param forceApplyAlignment If true, the applyAlignment value of the style will be set to true (used to merge cells)
     */
    public void setForceApplyAlignment(boolean forceApplyAlignment) {
        this.forceApplyAlignment = forceApplyAlignment;
    }

// ### C O N S T R U C T O R S ###   
     /**
      * Default constructor
      */
     public CellXf()
     {
         this.horizontalAlign = HorizontalAlignValue.none;
         this.alignment = TextBreakValue.none;
         this.textDirection = TextDirectionValue.horizontal;
         this.verticalAlign = VerticalAlignValue.none;
         this.textRotation = 0;            
     }

// ### M E T H O D S ###
     /**
      * Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
      * @return Returns the valid rotation in degrees for internal uses (LowLevel)
      * @throws RangeException Thrown if the rotation is out of range
      */
     public int calculateInternalRotation()
     {
         if (this.textRotation < -90 || this.textRotation > 90)
         {
             throw new RangeException("RotationRangeException","The rotation value (" + Integer.toString(this.textRotation) + "°) is out of range. Range is form -90° to +90°");
         }
         if (this.textDirection == TextDirectionValue.vertical)
         {
             return 255;
         }
         else
         {
             if (this.textRotation >= 0)
             {
                 return this.textRotation;
             }
             else
             {
                 return (90 - this.textRotation);
             }
         }            
     }

     /**
     * Override toString method
     * @return String of a class instance
     */
    @Override
    public String toString()
    {
        return this.getHash();
    }

    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */     
     @Override
     public CellXf copy()
     {
         CellXf copy = new CellXf();
         copy.setHorizontalAlign(this.horizontalAlign);
         copy.setAlignment(this.alignment);
                copy.setForceApplyAlignment(this.forceApplyAlignment);
                copy.setLocked(this.locked);
                copy.setHidden(this.hidden);         
         try
         {
         copy.setTextDirection(this.textDirection);
         copy.setTextRotation(this.textRotation);
         }
         catch (Exception e)
         {
             // Should never happen. Error will be thrown earlier on setting rotation in this instance
         }
         copy.setVerticalAlign(this.verticalAlign);
         return copy;
     } 
     
    /**
     * Override method to calculate the hash of this component
     * @return Calculated hash as string
     */     
    @Override
    String calculateHash() {
        StringBuilder sb = new StringBuilder();
        sb.append(CELLXFPREFIX);
        castValue(this.horizontalAlign.getValue(), sb, ':');
        castValue(this.verticalAlign, sb, ':');
        castValue(this.alignment.getValue(), sb, ':');
        castValue(this.textDirection.getValue(), sb, ':');
        castValue(this.textRotation, sb, ':');
        castValue(this.forceApplyAlignment, sb, ':');
        castValue(this.locked, sb, ':');
        castValue(this.hidden, sb, null);
        return sb.toString();
    }


}  
