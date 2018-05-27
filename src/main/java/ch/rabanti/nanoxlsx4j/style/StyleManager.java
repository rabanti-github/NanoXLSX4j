/*
 * NanoXLSX4j is a small Java library to write and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.nanoxlsx4j.style;

import ch.rabanti.nanoxlsx4j.exception.StyleException;

import java.util.ArrayList;
import java.util.Collections;

/**
 * Class representing a style manager to maintain all styles and its components of a workbook
 * @author Raphael Stockeli
 */
public class StyleManager
{
    
// ### C O N S T A N T S ###
    static final String BORDERPREFIX = "borders@";
    static final String CELLXFPREFIX = "/cellXf@";
    static final String FILLPREFIX = "/fill@";
    static final String FONTPREFIX = "/font@";
    static final String NUMBERFORMATPREFIX = "/numberFormat@";
    static final String STYLEPREFIX = "style=";
     
// ### P R I V A T E  F I E L D S ###
    private final ArrayList<AbstractStyle> borders;
    private final ArrayList<AbstractStyle> cellXfs;
    private final ArrayList<AbstractStyle> fills;
    private final ArrayList<AbstractStyle> fonts;
    private final ArrayList<AbstractStyle> numberFormats;
    private final ArrayList<AbstractStyle> styles;
    private final ArrayList<String> styleNames;
    
// ### C O N S T R U C T O R S ### 
    /**
     * Default constructor
     */
    public StyleManager()
    {
        this.borders = new ArrayList<>();
        this.cellXfs = new ArrayList<>();
        this.fills = new ArrayList<>();
        this.fonts = new ArrayList<>();
        this.numberFormats = new ArrayList<>();
        this.styles = new ArrayList<>();
        this.styleNames = new ArrayList<>();
    }
    
// ###  M E T H O D S ###
    /**
     * Gets a component by its hash
     * @param list List to check
     * @param hash Hash of the component
     * @return Determined component. If not found, null will be returned
     */
    private AbstractStyle getComponentByHash(ArrayList<AbstractStyle> list, String hash)
    {
        int len = list.size();
        for(int i = 0; i < len; i++)
        {
           if (list.get(i).getHash().equals(hash))
           {
               return list.get(i);
           }
        }
        return null;
    }
    
    /**
     * Gets a border by its hash
     * @param hash Hash of the border
     * @return Determined border
     * @throws StyleException Throws a StyleException if the border was not found in the style manager
     */
    public Border getBorderByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.borders, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style component with the hash '" + hash + "' was not found");
        }
        return (Border)component;
    }
    
    /**
     * Gets all borders of the style manager
     * @return Array of borders
     */
    public Border[] getBorders()
    {
        return this.borders.toArray(new Border[this.borders.size()]);
    }
    
    /**
     * Gets the number of borders in the style manager
     * @return Number of stored borders
     */
    public int getBorderStyleNumber()
    {
        return this.borders.size();
    }

    /**
     * Gets a cellXf by its hash
     * @param hash Hash of the cellXf
     * @return Determined cellXf
     * @throws StyleException Throws a StyleException if the cellXf was not found in the style manager
     */
    public CellXf getCellXfByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.cellXfs, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style component with the hash '" + hash + "' was not found");
        }
        return (CellXf)component;
    }
    
    /**
     * Gets all cellXfs of the style manager
     * @return Array of cellXfs
     */
    public CellXf[] getCellXfs()
    {
        return this.cellXfs.toArray(new CellXf[this.cellXfs.size()]);
    }   
    
     /**
     * Gets the number of cellXfs in the style manager
     * @return Number of stored cellXfs
     */
    public int getCellXfStyleNumber()
    {
        return this.cellXfs.size();
    }
    
    /**
     * Gets a font by its hash
     * @param hash Hash of the font
     * @return Determined font
     * @throws StyleException Throws a StyleException if the font was not found in the style manager
     */
    public Fill getFillByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.fills, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style component with the hash '" + hash + "' was not found");
        }
        return (Fill)component;
    }
    
    /**
     * Gets all fills of the style manager
     * @return Array of fills
     */
    public Fill[] getFills()
    {
        return this.fills.toArray(new Fill[this.fills.size()]);
    } 
    
     /**
     * Gets the number of fills in the style manager
     * @return Number of stored fills
     */
    public int getFillStyleNumber()
    {
        return this.fills.size();
    }    
    
    /**
     * Gets a font by its hash
     * @param hash Hash of the font
     * @return Determined font
     * @throws StyleException Throws a StyleException if the font was not found in the style manager
     */
    public Font getFontByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.fonts, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style component with the hash '" + hash + "' was not found");
        }
        return (Font)component;
    }
    
    /**
     * Gets all fonts of the style manager
     * @return Array of fonts
     */
    public Font[] getFonts()
    {
        return this.fonts.toArray(new Font[this.fonts.size()]);
    }  
    
     /**
     * Gets the number of fonts in the style manager
     * @return Number of stored fonts
     */
    public int getFontStyleNumber()
    {
        return this.fonts.size();
    }
    
    /**
     * Gets a number format by its hash
     * @param hash Hash of the number format
     * @return Determined number format
     * @throws StyleException Throws a StyleException if the number format was not found in the style manager
     */
    public NumberFormat getNumberFormatByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.numberFormats, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style component with the hash '" + hash + "' was not found");
        }
        return (NumberFormat)component;
    }
    
    /**
     * Gets all number formats of the style manager
     * @return Array of number formats
     */
    public NumberFormat[] getNumberFormats()
    {
        return this.numberFormats.toArray(new NumberFormat[this.numberFormats.size()]);
    }  
    
     /**
     * Gets the number of number formats in the style manager
     * @return Number of stored number formats
     */
    public int getNumberFormatStyleNumber()
    {
        return this.numberFormats.size();
    }
    
    /**
     * Gets a style by its name
     * @param name Name of the style
     * @return Determined style
     * @throws StyleException Throws a StyleException if the style was not found in the style manager
     */
    public Style getStyleByName(String name)
    {
        int len = this.styles.size();
        for(int i = 0; i < len; i++)
        {
           if (((Style)this.styles.get(i)).getName().equals(name))
           {
               return (Style)this.styles.get(i);
           }
        }
        throw new StyleException("MissingReferenceException","The style with the name '" + name + "' was not found");
    } 
    
    /**
     * Gets a style by its hash
     * @param hash Hash of the style
     * @return Determined style
     * @throws StyleException Throws a StyleException if the style was not found in the style manager
     */
    public Style getStyleByHash(String hash)
    {
        AbstractStyle component = getComponentByHash(this.styles, hash);
        if (component == null)
        {
            throw new StyleException("MissingReferenceException","The style with the hash '" + hash + "' was not found");
        }
        return (Style)component;
    }
    
    /**
     * Gets all styles of the style manager
     * @return Array of styles
     */
    public Style[] getStyles()
    {
        return this.styles.toArray(new Style[this.styles.size()]);

    } 
    
     /**
     * Gets the number of styles in the style manager
     * @return Number of stored styles
     */
    public int getStyleNumber()
    {
        return this.styles.size();
    }
    
    /**
     * Adds a style component to the manager
     * @param style Style to add
     * @return Added or determined style in the manager
     */
    public Style addStyle(Style style)
    {
        String hash = addStyleComponent(style);
        return (Style)this.getComponentByHash(this.styles, hash);
    }        
    
    /**
     * Adds a style component to the manager with an ID
     * @param style Component to add
     * @param id Id of the component
     * @return Hash of the added or determined component
     */
    private String addStyleComponent(AbstractStyle style, Integer id)
    {
        style.setInternalID(id);
        return addStyleComponent(style);
    }

    /**
     * Adds a style component to the manager
     * @param style Component to add
     * @return Hash of the added or determined component
     */
    private String addStyleComponent(AbstractStyle style)
    {
        String hash = style.getHash();
        if (style instanceof Border)
        {
            if (this.getComponentByHash(this.borders, hash) == null) { this.borders.add(style); }
            reorganize(borders);
        }
        else if (style instanceof CellXf)
        {
            if (this.getComponentByHash(this.cellXfs, hash) == null) { this.cellXfs.add(style); }
            reorganize(cellXfs);
        }
        else if (style instanceof Fill)
        {
            if (this.getComponentByHash(this.fills, hash) == null) { this.fills.add(style); }
            reorganize(fills);
        }
        else if (style instanceof Font)
        {
            if (this.getComponentByHash(this.fonts, hash) == null) { this.fonts.add(style); }
            reorganize(fonts);
        }
        else if (style instanceof NumberFormat)
        {
            if (this.getComponentByHash(this.numberFormats, hash) == null) { this.numberFormats.add(style); }
            reorganize(numberFormats);
        }
        else if (style instanceof Style)
        {
            Style s = (Style)style;
            if (this.styleNames.contains(s.getName()) == true)
            {
                throw new StyleException("StyleAlreadyExistsException","The style with the name '" + s.getName() + "' already exists");
            }
            if (this.getComponentByHash(this.styles, hash) == null)
            {
                Integer id;
                if (s.getInternalID() == null)
                {
                    id = Integer.MAX_VALUE;
                    s.setInternalID(id);
                }
                else
                {
                    id = s.getInternalID();
                }
                String temp = this.addStyleComponent(s.getBorder(), id);
                s.setBorder((Border)this.getComponentByHash(this.borders, temp));
                temp = this.addStyleComponent(s.getCellXf(), id);
                s.setCellXf((CellXf)this.getComponentByHash(this.cellXfs, temp));
                temp = this.addStyleComponent(s.getFill(), id);
                s.setFill((Fill)this.getComponentByHash(this.fills, temp));
                temp = this.addStyleComponent(s.getFont(), id);
                s.setFont((Font)this.getComponentByHash(this.fonts, temp));
                temp = this.addStyleComponent(s.getNumberFormat(), id);
                s.setNumberFormat((NumberFormat)this.getComponentByHash(this.numberFormats, temp));
                this.styles.add(s);
            }
            reorganize(styles);
            hash = s.calculateHash();
        }
        return hash;
    }
    
    /**
     * Removes a style and all its components from the style manager
     * @param styleName Name of the style to remove
     * @throws StyleException Throws a StyleException if the style was not found in the style manager
     */
    public void removeStyle(String styleName)
    {
//        String hash = null;
        boolean match = false;
        int len = this.styles.size();
        int index = -1;
        for(int i = 0; i < len; i++)
        {
            if (((Style)this.styles.get(i)).getName().equals(styleName) == true)
            { 
                match = true;
//                hash = ((Style)this.styles.get(i)).getHash();
                index = i;
                break; 
            }
        }
        if (match == false)
        {
            throw new StyleException("MissingReferenceException","The style with the name '" + styleName + "' was not found in the style manager");
        }
        this.styles.remove(index);
        cleanupStyleComponents();
    }    
    
    /**
     * Method to reorganize / reorder a list of style components
     * @param list List to reorganize
     */
    private void reorganize(ArrayList<AbstractStyle> list)
    {
        int len = list.size();
        Collections.sort(list);
        int id = 0;
        for(int i = 0; i < len; i++)
        {
            list.get(i).setInternalID(id);
            id++;
        }
    }    
    
    /**
     * Method to cleanup style components in the style manager
     */
    private void cleanupStyleComponents()
    {
        Border border;
        CellXf cellXf;
        Fill fill;
        Font font;
        NumberFormat numberFormat;
        int len = this.borders.size();
        int i;
        for(i = len; i >= 0; i--)
        {
           border = (Border)this.borders.get(i);
           if (isUsedByStyle(border) == false) { this.borders.remove(i); }
        }
        len = this.cellXfs.size();
        for(i = len; i >= 0; i--)
        {
           cellXf = (CellXf)this.cellXfs.get(i);
           if (isUsedByStyle(cellXf) == false) { this.cellXfs.remove(i); }
        }
        len = this.fills.size();
        for(i = len; i >= 0; i--)
        {
           fill = (Fill)this.fills.get(i);
           if (isUsedByStyle(fill) == false) { this.fills.remove(i); }
        }
        len = this.fonts.size();
        for(i = len; i >= 0; i--)
        {
           font = (Font)this.fonts.get(i);
           if (isUsedByStyle(font) == false) { this.fonts.remove(i); }
        }
        len = this.numberFormats.size();
        for(i = len; i >= 0; i--)
        {
           numberFormat = (NumberFormat)this.numberFormats.get(i);
           if (isUsedByStyle(numberFormat) == false) { this.numberFormats.remove(i); }
        }
    }    

    /**
     * Checks whether a style component in the style manager is used by a style
     * @param component Component to check
     * @return If true, the component is in use
     */
    private boolean isUsedByStyle(AbstractStyle component)
    {
        Style s;
        boolean match = false;
        String hash = component.getHash();
        int len = this.styles.size();
        for(int i = 0; i < len; i++)
        {
            s = (Style)this.styles.get(i);
            if (component instanceof Border) { if (s.getBorder().getHash().equals(hash)) { match = true; break; } }
            else if (component instanceof CellXf) { if (s.getCellXf().getHash().equals(hash)) { match = true; break; } }
            if (component instanceof Fill) { if (s.getFill().getHash().equals(hash)) { match = true; break; } }
            if (component instanceof Font) { if (s.getFont().getHash().equals(hash)) { match = true; break; } }
            if (component instanceof NumberFormat) { if (s.getNumberFormat().getHash().equals(hash)) { match = true; break; } }
        }
        return match;
    }
}