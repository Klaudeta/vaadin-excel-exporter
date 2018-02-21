/*
 * 
 */
package org.vaadin.addons.excelexporter.configuration;

import java.awt.Color;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

import net.karneim.pojobuilder.GeneratePojoBuilder;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.vaadin.addons.excelexporter.formatter.ColumnFormatter;
import org.vaadin.addons.excelexporter.function.DataCellStyleGeneratorFunction;
import org.vaadin.addons.excelexporter.utils.ExcelStyleUtil;

import com.vaadin.ui.Grid;

/**
 * The Class ExportExcelComponentConfiguration is used to configure the
 * component as in Table, Tree Table or Grid Several additional properties such
 * as formatting, headers, footers, styles etc. can also be configured
 *
 * @author Kartik Suba
 */
@GeneratePojoBuilder(intoPackage = "*.builder")
public class ExportExcelComponentConfiguration<BEANTYPE> {

	/** The visible properties. */
	private String[] visibleProperties;

	/** The date formatting properties. */
	private List<String> dateFormattingProperties;

	/** The integer formatting properties. */
	private List<String> integerFormattingProperties;

	/** The float formatting properties. */
	private List<String> floatFormattingProperties;

	/** The boolean formatting properties. */
	private List<String> booleanFormattingProperties;

	/** The header configs. */
	private List<ComponentHeaderConfiguration> headerConfigs;

	/** The footer configs. */
	private List<ComponentFooterConfiguration> footerConfigs;

	/** The table header style function. */
  private BiFunction<Workbook, String, CellStyle> headerStyleFunction = (workbook, columnId) -> {
    CellStyle headerCellStyle = workbook.createCellStyle();

    headerCellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
    headerCellStyle.setAlignment(CellStyle.ALIGN_LEFT);

    if (headerCellStyle instanceof XSSFCellStyle) {
      ((XSSFCellStyle) headerCellStyle).setFillForegroundColor(new XSSFColor(new Color(50, 86, 110)));
      headerCellStyle = ExcelStyleUtil.setBorders((XSSFCellStyle) headerCellStyle, Boolean.TRUE, Boolean.TRUE,
          Boolean.TRUE, Boolean.TRUE, Color.WHITE);
    }
    // XSSFFont boldFont = workbook.createFont();
    Font boldFont = workbook.createFont();
    boldFont.setColor(IndexedColors.WHITE.getIndex());
    boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);

    headerCellStyle.setFont(boldFont);

    return headerCellStyle;
  };

	/** The table footer style function. */
	private BiFunction<Workbook, String, CellStyle> footerStyleFunction = (workbook, columnId) -> {
		CellStyle headerCellStyle = workbook.createCellStyle();

		headerCellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		headerCellStyle.setAlignment(CellStyle.ALIGN_LEFT);

		if(headerCellStyle instanceof XSSFCellStyle) {
		  
	    ((XSSFCellStyle) headerCellStyle).setFillForegroundColor(new XSSFColor(new Color(50, 86, 110)));
	    headerCellStyle = ExcelStyleUtil.setBorders(((XSSFCellStyle) headerCellStyle), Boolean.TRUE, Boolean.TRUE, Boolean.TRUE,
          Boolean.TRUE, Color.WHITE);
		}
		
		Font boldFont = workbook.createFont();
		boldFont.setColor(IndexedColors.WHITE.getIndex());
		boldFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		
		headerCellStyle.setFont(boldFont);

		return headerCellStyle;
	};

  /** The table content style function. */
  private DataCellStyleGeneratorFunction contentStyleFunction = new DefaultContentStyleFunction();

	/** The column formatters. */
	private Map<Object, ColumnFormatter> columnFormatters = new LinkedHashMap<>();

	/** The grid. */
	private Grid<BEANTYPE> grid;

	/**
	 * Gets the grid.
	 *
	 * @return the grid
	 */
	public Grid<BEANTYPE> getGrid() {
		return this.grid;
	}

	/**
	 * Sets the grid.
	 *
	 * @param grid
	 *            the new grid
	 */
	public void setGrid(final Grid<BEANTYPE> grid) {
		this.grid = grid;
	}

	public BiFunction<Workbook, String, CellStyle> getHeaderStyleFunction() {
		return this.headerStyleFunction;
	}

	public void setHeaderStyleFunction(BiFunction<Workbook, String, CellStyle> headerStyleFunction) {
		this.headerStyleFunction = headerStyleFunction;
	}

	public BiFunction<Workbook, String, CellStyle> getFooterStyleFunction() {
		return this.footerStyleFunction;
	}

	public void setFooterStyleFunction(BiFunction<Workbook, String, CellStyle> footerStyleFunction) {
		this.footerStyleFunction = footerStyleFunction;
	}

	public DataCellStyleGeneratorFunction getContentStyleFunction() {
		return this.contentStyleFunction;
	}

	public void setContentStyleFunction(DataCellStyleGeneratorFunction contentStyleFunction) {
		this.contentStyleFunction = contentStyleFunction;
	}

	/**
	 * Gets the date formatting properties.
	 *
	 * @return the date formatting properties
	 */
	public List<String> getDateFormattingProperties() {
		if (this.dateFormattingProperties == null) {
			this.dateFormattingProperties = new ArrayList<>();
		}
		return this.dateFormattingProperties;
	}

	/**
	 * Sets the date formatting properties.
	 *
	 * @param dateFormattingProperties
	 *            the new date formatting properties
	 */
	public void setDateFormattingProperties(final List<String> dateFormattingProperties) {
		this.dateFormattingProperties = dateFormattingProperties;
	}

	/**
	 * Gets the integer formatting properties.
	 *
	 * @return the integer formatting properties
	 */
	public List<String> getIntegerFormattingProperties() {
		if (this.integerFormattingProperties == null) {
			this.integerFormattingProperties = new ArrayList<>();
		}
		return this.integerFormattingProperties;
	}

	/**
	 * Sets the integer formatting properties.
	 *
	 * @param integerFormattingProperties
	 *            the new integer formatting properties
	 */
	public void setIntegerFormattingProperties(final List<String> integerFormattingProperties) {
		this.integerFormattingProperties = integerFormattingProperties;
	}

	/**
	 * Gets the float formatting properties.
	 *
	 * @return the float formatting properties
	 */
	public List<String> getFloatFormattingProperties() {
		if (this.floatFormattingProperties == null) {
			this.floatFormattingProperties = new ArrayList<>();
		}
		return this.floatFormattingProperties;
	}

	/**
	 * Sets the float formatting properties.
	 *
	 * @param floatFormattingProperties
	 *            the new float formatting properties
	 */
	public void setFloatFormattingProperties(final List<String> floatFormattingProperties) {
		this.floatFormattingProperties = floatFormattingProperties;
	}

	/**
	 * Sets the boolean formatting properties.
	 *
	 * @param booleanFormattingProperties
	 *            the new boolean formatting properties
	 */
	public void setBooleanFormattingProperties(final List<String> booleanFormattingProperties) {
		this.booleanFormattingProperties = booleanFormattingProperties;
	}

	/**
	 * Sets the header configs.
	 *
	 * @param headerConfigs
	 *            the new header configs
	 */
	public void setHeaderConfigs(final List<ComponentHeaderConfiguration> headerConfigs) {
		this.headerConfigs = headerConfigs;
	}

	/**
	 * Sets the footer configs.
	 *
	 * @param footerConfigs
	 *            the new footer configs
	 */
	public void setFooterConfigs(final List<ComponentFooterConfiguration> footerConfigs) {
		this.footerConfigs = footerConfigs;
	}

	/**
	 * Sets the column formatters.
	 *
	 * @param columnFormatters
	 *            the column formatters
	 */
	public void setColumnFormatters(final Map<Object, ColumnFormatter> columnFormatters) {
		this.columnFormatters = columnFormatters;
	}

	/**
	 * Gets the visible properties.
	 *
	 * @return the visible properties
	 */
	public String[] getVisibleProperties() {
		return this.visibleProperties;
	}

	/**
	 * Sets the visible properties.
	 *
	 * @param visibleProperties
	 *            the new visible properties
	 */
	public void setVisibleProperties(final String[] visibleProperties) {
		this.visibleProperties = visibleProperties;
	}

	/**
	 * Gets the boolean formatting properties.
	 *
	 * @return the boolean formatting properties
	 */
	public List<String> getBooleanFormattingProperties() {
		if (this.booleanFormattingProperties == null) {
			this.booleanFormattingProperties = new ArrayList<>();
		}
		return this.booleanFormattingProperties;
	}

	/**
	 * Gets the header configs.
	 *
	 * @return the header configs
	 */
	public List<ComponentHeaderConfiguration> getHeaderConfigs() {
		return this.headerConfigs;
	}

	/**
	 * Gets the footer configs.
	 *
	 * @return the footer configs
	 */
	public List<ComponentFooterConfiguration> getFooterConfigs() {
		return this.footerConfigs;
	}

	/**
	 * Gets the column formatters.
	 *
	 * @return the column formatters
	 */
	public Map<Object, ColumnFormatter> getColumnFormatters() {
		return this.columnFormatters;
	}

	/**
	 * Gets the column formatter.
	 *
	 * @param columnId
	 *            the column id
	 * @return the column formatter
	 */
	public ColumnFormatter getColumnFormatter(final Object columnId) {
		if (this.columnFormatters != null && !this.columnFormatters.isEmpty()
				&& this.columnFormatters.containsKey(columnId)) {
			return this.columnFormatters.get(columnId);
		}
		return null;
	}
}

class DefaultContentStyleFunction implements DataCellStyleGeneratorFunction {

  private final CacheableValue<CellStyle> dataStyle1 = new CacheableValue<>();
  private final CacheableValue<CellStyle> dataStyle2 = new CacheableValue<>();

  @Override
  public CellStyle apply(Workbook workbook, String columnId, Object value, int rowNum) {
    if (rowNum % 2 == 0) {
      return dataStyle1.computeIfAbsent(workbook::createCellStyle);
    } else {
      return dataStyle2.computeIfAbsent(() -> {
        CellStyle style = workbook.createCellStyle();
        // XXX Only supported on XSSFCellStyle
        // dataStyle2.setFillForegroundColor(new XSSFColor(new Color(228, 234,
        // 238)));
        // TODO
        style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        return style;
      });
    }
  }
}