/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package xlsreport.report;

import com.mendix.core.Core;
import com.mendix.logging.ILogNode;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import xlsreport.proxies.*;
import xlsreport.report.export.Export;

import java.util.HashMap;
import java.util.List;

/**
 *
 * @author jvg
 */
public class Styling
{
	private static final String DATEFORMAT = "_Date";
	private static ILogNode log = Core.getLogger("XLSreport");   
	private HashMap<String, CellStyle> styleList;
    private CellStyle defaultStyle;
    private CellStyle defaultStyleDate;
    private CreationHelper createHelper;
    private String datePresentation;
    
    public Styling(MxTemplate template)
    {
    	this.datePresentation = Export.getDatePresentation(template);
    	this.styleList = new HashMap<String, CellStyle>();
    }

    public CellStyle getDefaultStyle()
    {
        return this.defaultStyle;
    }
    
    public CellStyle getDefaultStyle(boolean dateTimeFormat)
    {
    	if(dateTimeFormat)
    	{
    		return this.defaultStyleDate;
    	} else
    	{
    		return this.defaultStyle;
    	}
    }

    public CellStyle getStyle(Long name, boolean useDateTimeFormat)
    {
    	if(useDateTimeFormat)
    	{    		
    		return styleList.get(name.toString()+DATEFORMAT);
    	} else
    	{
    		return styleList.get(name.toString());
    	}
    }

    public void setAllStyles(IContext context, MxTemplate mxTemplate, Workbook book)
    {
        if (styleList != null)
        {
            styleList.clear();
        }
        this.createHelper = book.getCreationHelper();
        //CreationHelper createHelper = book.getCreationHelper();
        log.debug("-- Initialise all the styles to the hashmap.");
        List<IMendixObject> stylesObjects = Core.createXPathQuery("//" + MxCellStyle.getType()
                        + "[XLSReport.MxCellStyle_Template='" + mxTemplate.getMendixObject().getId().toLong() + "']")
                .execute(context);
        for (IMendixObject styleObject : stylesObjects)
        {
            // Create a style
            MxCellStyle MxStyle = MxCellStyle.initialize(context, styleObject);            
            
            // Create a normal version of the style
            CellStyle style = createCellStyle(MxStyle, book, false);
            this.styleList.put(MxStyle.getMendixObject().getId().toLong()+"", style);
            
            // Create a normal version of the style
            var dateStyle = createCellStyle(MxStyle, book, true);
            this.styleList.put(MxStyle.getMendixObject().getId().toLong()+DATEFORMAT, dateStyle);
        }
    }

    public CellStyle createCellStyle(MxCellStyle MxStyle, Workbook book, boolean dateTimeFormat)
    {
    	CellStyle style = book.createCellStyle();
        // Create the font for the style.
        Font font = book.createFont();
        font.setItalic(MxStyle.getTextItalic());            
        font.setFontHeightInPoints(MxStyle.getTextHeight().shortValue());
        if (MxStyle.getTextColor() != null && MxStyle.getTextColor() != MxColor.Blank)
        {
            font.setColor(getColor(MxStyle.getTextColor()));
        }
        if (MxStyle.getTextUnderline())
        {
            font.setUnderline(Font.U_SINGLE);
        }
        if (MxStyle.getTextBold())
        {
            font.setBold(true);
        }
        style.setFont(font);
        // Alignment
        style.setAlignment(getAlignment(MxStyle.getTextAlignment()));
        style.setVerticalAlignment(getVerticalAlignment(MxStyle.getTextVerticalalignment()));
        // Color fill and other options.
        if (MxStyle.getBackgroundColor() != null && MxStyle.getBackgroundColor() != MxColor.Blank)
        {
            style.setFillForegroundColor(getColor(MxStyle.getBackgroundColor()));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        style.setRotation(MxStyle.getTextRotation().shortValue());
        if (MxStyle.getTextRotation() == 0)
        {
            style.setWrapText(MxStyle.getWrapText());
        }
        // Create border lines.
        if (MxStyle.getBorderTop() > 0)
        {
            style.setBorderTop(BorderStyle.valueOf(MxStyle.getBorderTop().shortValue()));
        }
        if (MxStyle.getBorderBottom() > 0)
        {
            style.setBorderBottom(BorderStyle.valueOf(MxStyle.getBorderBottom().shortValue()));
        }
        if (MxStyle.getBorderLeft() > 0)
        {
            style.setBorderLeft(BorderStyle.valueOf(MxStyle.getBorderLeft().shortValue()));
        }
        if (MxStyle.getBorderRight() > 0)
        {
            style.setBorderRight(BorderStyle.valueOf(MxStyle.getBorderRight().shortValue()));
        }
        if (MxStyle.getBorderColor() != null && MxStyle.getBorderColor() != MxColor.Blank)
        {
            style.setTopBorderColor(getColor(MxStyle.getBorderColor()));
            style.setBottomBorderColor(getColor(MxStyle.getBorderColor()));
            style.setLeftBorderColor(getColor(MxStyle.getBorderColor()));
            style.setRightBorderColor(getColor(MxStyle.getBorderColor()));
        }

        if(MxStyle.getDataFormatString() != null)
        {
            style.setDataFormat(this.createHelper.createDataFormat().getFormat(MxStyle.getDataFormatString()));
        }

        if(dateTimeFormat)
        {
          style.setDataFormat(this.createHelper.createDataFormat().getFormat(this.datePresentation));
        	log.trace("Created style with DateTimeFormat: " + style.getDataFormatString());
        }
    	return style;
    }
    
    private static short getColor(MxColor color)
    {
        switch (color)
        {
            case Black:
                return HSSFColor.HSSFColorPredefined.BLACK.getIndex();
            case Blue:
                return HSSFColor.HSSFColorPredefined.BLUE.getIndex();
            case Brown:
                return HSSFColor.HSSFColorPredefined.BROWN.getIndex();
            case Green:
                return HSSFColor.HSSFColorPredefined.GREEN.getIndex();
            case Light_Blue:
                return HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex();
            case Orange:
                return HSSFColor.HSSFColorPredefined.ORANGE.getIndex();
            case Pink:
                return HSSFColor.HSSFColorPredefined.PINK.getIndex();
            case Red:
                return HSSFColor.HSSFColorPredefined.RED.getIndex();
            case White:
                return HSSFColor.HSSFColorPredefined.WHITE.getIndex();
            case Yellow:
                return HSSFColor.HSSFColorPredefined.YELLOW.getIndex();
            case Gray_1:
                return HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex();
            case Gray_2:
                return HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex();
            case Gray_3:
                return HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getIndex();
            case Gray_4:
                return HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.getIndex();
            default:
                return HSSFColor.HSSFColorPredefined.WHITE.getIndex();
        }
    }

    private static HorizontalAlignment getAlignment(TextAlignment align)
    {
        if (align != null)
        {
            switch (align)
            {
                case Left:
                    return HorizontalAlignment.LEFT;
                case Center:
                    return HorizontalAlignment.CENTER;
                case Right:
                    return HorizontalAlignment.RIGHT;
            }
        }       
        return HorizontalAlignment.LEFT;
    }

    private static VerticalAlignment getVerticalAlignment(TextVerticalAlignment align)
    {
        if (align != null)
        {
            switch (align)
            {
                case Top:
                    return VerticalAlignment.TOP;
                case Middle:
                    return VerticalAlignment.CENTER;
                case Bottom:
                    return VerticalAlignment.BOTTOM;
            }
        }
        return VerticalAlignment.TOP;
    }
}
