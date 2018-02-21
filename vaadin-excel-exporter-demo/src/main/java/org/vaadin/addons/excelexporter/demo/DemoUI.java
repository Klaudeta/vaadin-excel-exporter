package org.vaadin.addons.excelexporter.demo;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.annotation.WebServlet;

import org.vaadin.addons.excelexporter.ExportToExcel;
import org.vaadin.addons.excelexporter.configuration.ComponentFooterConfiguration;
import org.vaadin.addons.excelexporter.configuration.ComponentHeaderConfiguration;
import org.vaadin.addons.excelexporter.configuration.ExportExcelComponentConfiguration;
import org.vaadin.addons.excelexporter.configuration.ExportExcelConfiguration;
import org.vaadin.addons.excelexporter.configuration.ExportExcelSheetConfiguration;
import org.vaadin.addons.excelexporter.configuration.MergedCell;
import org.vaadin.addons.excelexporter.demo.data.DataModelGenerator;
import org.vaadin.addons.excelexporter.formatter.BooleanColumnFormatter;
import org.vaadin.addons.excelexporter.formatter.ColumnFormatter;
import org.vaadin.addons.excelexporter.formatter.SuffixColumnFormatter;
import org.vaadin.addons.excelexporter.model.ExportType;

import com.vaadin.annotations.Title;
import com.vaadin.annotations.VaadinServletConfiguration;
import com.vaadin.annotations.Widgetset;
import com.vaadin.data.provider.CallbackDataProvider;
import com.vaadin.data.provider.ListDataProvider;
import com.vaadin.icons.VaadinIcons;
import com.vaadin.server.VaadinRequest;
import com.vaadin.server.VaadinServlet;
import com.vaadin.ui.Grid;
import com.vaadin.ui.MenuBar;
import com.vaadin.ui.TabSheet;
import com.vaadin.ui.UI;
import com.vaadin.ui.VerticalLayout;
import com.vaadin.ui.components.grid.FooterCell;
import com.vaadin.ui.components.grid.FooterRow;
import com.vaadin.ui.components.grid.HeaderCell;
import com.vaadin.ui.components.grid.HeaderRow;

/**
 *
 * @author Kartik Suba
 *
 */
@Title("vaadin-excel-exporter Add-on Demo")
@SuppressWarnings("serial")
@Widgetset("AppWidgetset")
public class DemoUI extends UI {

  private Grid<DataModel> gridDefault;
  private Grid<DataModel> gridMergedCells;
  private Grid<DataModel> gridFrozenColumns;

  private transient String[] visibleColumns = new String[] { "country", "productType", "catalogue", "plannedPrinter",
      "cheapest", "contractor", "totalCosts", "differenceToMin", "comment", "active", "counter" };
  private transient String[] columnHeaders = new String[] { "COUNTRY", "PRODUCT_TYPE", "CATALOGUE", "PLANNED_PRINTER",
      "CHEAPEST", "CONTRACTOR", "TOTAL_COST", "DIFFERENCE_TO_MIN", "COMMENT", "ACTIVE", "COUNTER" };

  @WebServlet(value = "/*", asyncSupported = true)
  @VaadinServletConfiguration(productionMode = false, ui = DemoUI.class, widgetset = "org.vaadin.addons.demo.DemoWidgetSet")
  public static class Servlet extends VaadinServlet {
  }

  @Override
  protected void init(final VaadinRequest request) {

    // Creating the Export Tool Bar
    MenuBar exportToolBar = createToolBar();

    final VerticalLayout layout = new VerticalLayout();
    layout.setSizeFull();

    // Adding the Export Tool Bar to the Layout
    layout.addComponent(exportToolBar);

    /*********
     * Adding Components to the Layout namely Tables, Grids and Tree Table
     *******/
    //Invalid row number (1048576) outside allowable range (0..1048575)
    final CallbackDataProvider<DataModel, String> dataProvider = new CallbackDataProvider<>(
        q -> DataModelGenerator.generate(q.getLimit()).stream(), q -> 1000 * 1000);
    this.gridDefault = new Grid<>(DataModel.class);
    this.gridDefault.setDataProvider(dataProvider);
    this.gridDefault.setSizeFull();
    this.gridDefault.setColumns(this.visibleColumns);

    this.gridMergedCells = new Grid<>(DataModel.class);
    this.gridMergedCells.setDataProvider(new ListDataProvider<>(DataModelGenerator.generate(20)));
    this.gridMergedCells.setColumns(this.visibleColumns);
    this.gridMergedCells.setSizeFull();
    HeaderRow headerRow = this.gridMergedCells.addHeaderRowAt(0);
    HeaderCell joinHeaderColumns1 = headerRow.join("country", "productType");
    joinHeaderColumns1.setText("mergedCell");
    HeaderCell joinHeaderColumns2 = headerRow.join("cheapest", "contractor");
    joinHeaderColumns2.setText("mergedCell");
    FooterRow footerRow1 = this.gridMergedCells.addFooterRowAt(0);
    FooterCell joinFooterColumns1 = footerRow1.join("country", "productType");
    joinFooterColumns1.setText("mergedCell");
    FooterCell joinFooterColumns2 = footerRow1.join("cheapest", "contractor");
    joinFooterColumns2.setText("mergedCell");
    FooterRow footerRow2 = this.gridMergedCells.addFooterRowAt(0);
    for (int i = 0; i < this.visibleColumns.length; i++) {
      footerRow2.getCell(this.visibleColumns[i]).setText(this.columnHeaders[i]);
    }

    this.gridFrozenColumns = new Grid<>(DataModel.class);
    this.gridFrozenColumns.setDataProvider(new ListDataProvider<>(DataModelGenerator.generate(20)));
    this.gridFrozenColumns.setColumns(this.visibleColumns);
    this.gridFrozenColumns.getColumn("country").setWidth(300);
    this.gridFrozenColumns.getColumn("productType").setWidth(300);
    this.gridFrozenColumns.getColumn("catalogue").setWidth(300);
    this.gridFrozenColumns.setSizeFull();
    this.gridFrozenColumns.setFrozenColumnCount(3);

    TabSheet tabSheet = new TabSheet();
    tabSheet.setSizeFull();
    tabSheet.addTab(this.gridDefault, "Grid (Default)");
    tabSheet.addTab(this.gridMergedCells, "Grid (Merged Cells)");
    tabSheet.addTab(this.gridFrozenColumns, "Grid (Frozen Columns&Rows)");
    layout.addComponent(tabSheet);
    layout.setExpandRatio(tabSheet, 1);

    /*********
     * Adding Components to the Layout namely Tables, Grids and Tree Table
     *******/

    /*********
     * Adding the above data to the containers or components
     *******/

    setContent(layout);
  }

  /**
   * This method creates the tool bar with options XLSX and XLS.
   * 
   * @return
   */
  private MenuBar createToolBar() {
    MenuBar exportToolBar = new MenuBar();

    exportToolBar.addItem("XLSX", VaadinIcons.DOWNLOAD, selectedItem -> {
      ExportToExcel<DataModel> exportToExcelUtility = customizeExportExcelUtility(ExportType.XLSX);
      exportToExcelUtility.export();
    });

    exportToolBar.addItem("XLS", VaadinIcons.DOWNLOAD, selectedItem -> {
      ExportToExcel<DataModel> exportToExcelUtility = customizeExportExcelUtility(ExportType.XLS);
      exportToExcelUtility.export();
    });

    return exportToolBar;
  }

  /**
   * Configuring ExportToExcel Utility This configuration allows the end
   * user-developer to \ Add multiple sheets and configure them separately \
   * Configure components to be added in each sheet and their properties
   * 
   * @param exportType
   *
   * @return ExportToExcelUtility
   */
  private ExportToExcel<DataModel> customizeExportExcelUtility(ExportType exportType) {

    HashMap<Object, ColumnFormatter> columnFormatters = new HashMap<>();

    // Suffix Formatter provided
    columnFormatters.put("totalCosts", new SuffixColumnFormatter("$"));
    columnFormatters.put("differenceToMin", new SuffixColumnFormatter("$"));
    columnFormatters.put("cheapest", new SuffixColumnFormatter("-quite cheap"));

    // Boolean Formatter provided
    columnFormatters.put("active", new BooleanColumnFormatter("Yes", "No"));

    // Custom Formatting also possible
    columnFormatters.put("catalogue", (final Object value, final Object itemId,
        final Object columnId) -> (value != null) ? ((String) value).toLowerCase() : null);

    ExportExcelComponentConfiguration<DataModel> componentConfig1 = new ExportExcelComponentConfigurationBuilder<DataModel>()
        .withGrid(this.gridDefault).withVisibleProperties(this.visibleColumns)
        .withHeaderConfigs(Arrays.asList(
            new ComponentHeaderConfigurationBuilder().withAutoFilter(true).withColumnKeys(this.columnHeaders).build()))
        .withIntegerFormattingProperties(Arrays.asList("counter"))
        .withFloatFormattingProperties(Arrays.asList("totalCosts", "differenceToMin"))
        .withBooleanFormattingProperties(Arrays.asList("active")).withColumnFormatters(columnFormatters).build();

    ExportExcelComponentConfiguration<DataModel> componentConfig2 = new ExportExcelComponentConfigurationBuilder<DataModel>()
        .withGrid(
            this.gridMergedCells)
        .withVisibleProperties(this.visibleColumns)
        .withHeaderConfigs(
            Arrays
                .asList(
                    new ComponentHeaderConfigurationBuilder().withMergedCells(Arrays.asList(
                        new MergedCellBuilder().withStartProperty("country").withEndProperty("productType")
                            .withHeaderKey("mergedCell").build(),
                        new MergedCellBuilder().withStartProperty("cheapest").withEndProperty("contractor")
                            .withHeaderKey("mergedCell").build()))
                        .withColumnKeys(this.columnHeaders).build(),
                    new ComponentHeaderConfigurationBuilder().withAutoFilter(true).withColumnKeys(this.columnHeaders)
                        .build()))
        .withFooterConfigs(
            Arrays.asList(new ComponentFooterConfigurationBuilder().withColumnKeys(this.columnHeaders).build(),
                new ComponentFooterConfigurationBuilder().withMergedCells(Arrays.asList(
                    new MergedCellBuilder().withStartProperty("country").withEndProperty("productType")
                        .withHeaderKey("mergedCell").build(),
                    new MergedCellBuilder().withStartProperty("cheapest").withEndProperty("contractor")
                        .withHeaderKey("mergedCell").build()))
                    .withColumnKeys(this.columnHeaders).build()))
        .withIntegerFormattingProperties(Arrays.asList("counter"))
        .withFloatFormattingProperties(Arrays.asList("totalCosts", "differenceToMin"))
        .withBooleanFormattingProperties(Arrays.asList("active")).withColumnFormatters(columnFormatters).build();

    /* Configuring Sheets */
    ExportExcelSheetConfiguration<DataModel> sheetConfig1 = new ExportExcelSheetConfigurationBuilder<DataModel>()
        .withReportTitle("Grid (Default)").withSheetName("Grid (default)")
        .withComponentConfigs(Arrays.asList(componentConfig1)).withIsHeaderSectionRequired(Boolean.TRUE)
        .withDateFormat("dd-MMM-yyyy").build();

    ExportExcelSheetConfiguration<DataModel> sheetConfig2 = new ExportExcelSheetConfigurationBuilder<DataModel>()
        .withReportTitle("Grid (Merged Cells)").withSheetName("Grid (merged Cells)")
        .withComponentConfigs(Arrays.asList(componentConfig1)).withIsHeaderSectionRequired(Boolean.TRUE)
        .withDateFormat("dd-MMM-yyyy").build();

    ExportExcelSheetConfiguration<DataModel> sheetConfig3 = new ExportExcelSheetConfigurationBuilder<DataModel>()
        .withReportTitle("Grid (Frozen Columns&Rows)").withSheetName("Grid (Frozen Columns&Rows)").withFrozenColumns(3)
        .withFrozenRows(5).withComponentConfigs(Arrays.asList(componentConfig1))
        .withIsHeaderSectionRequired(Boolean.TRUE).withDateFormat("dd-MMM-yyyy").build();

    ExportExcelSheetConfiguration<DataModel> sheetConfig4 = new ExportExcelSheetConfigurationBuilder<DataModel>()
        .withReportTitle("Exported multiple Grids into one sheet").withSheetName("Multiple Grids")
        .withComponentConfigs(Arrays.asList(componentConfig1, componentConfig2))
        .withIsDefaultSheetTitleRequired(Boolean.FALSE).withIsHeaderSectionRequired(Boolean.FALSE)
        .withDateFormat("dd-MMM-yyyy").build();

    /* Configuring Excel */
    ExportExcelConfiguration<DataModel> config1 = new ExportExcelConfigurationBuilder<DataModel>()
        .withGeneratedBy("Kartik Suba")
        .withSheetConfigs(Arrays.asList(sheetConfig1, sheetConfig2, sheetConfig3, sheetConfig4)).build();

    return new ExportToExcel<>(exportType, config1);
  }

  class ComponentFooterConfigurationBuilder {
    private ComponentFooterConfiguration bean = new ComponentFooterConfiguration();

    ComponentFooterConfigurationBuilder withColumnKeys(String[] columnKeys) {
      bean.setColumnKeys(columnKeys);
      return this;
    }

    ComponentFooterConfigurationBuilder withMergedCells(List<MergedCell> mergedCells) {
      bean.setMergedCells(mergedCells);
      return this;
    }

    ComponentFooterConfiguration build() {
      return bean;
    }
  };

  class ComponentHeaderConfigurationBuilder {
    private ComponentHeaderConfiguration bean = new ComponentHeaderConfiguration();

    ComponentHeaderConfigurationBuilder withAutoFilter(boolean withAutoFilter) {
      bean.setAutoFilter(withAutoFilter);
      return this;
    }

    ComponentHeaderConfigurationBuilder withColumnKeys(String[] columnKeys) {
      bean.setColumnKeys(columnKeys);
      return this;
    }

    ComponentHeaderConfigurationBuilder withMergedCells(List<MergedCell> mergedCells) {
      bean.setMergedCells(mergedCells);
      return this;
    }

    ComponentHeaderConfiguration build() {
      return bean;
    }
  };

  class ExportExcelComponentConfigurationBuilder<T> {
    private ExportExcelComponentConfiguration<T> bean = new ExportExcelComponentConfiguration<T>();

    ExportExcelComponentConfigurationBuilder<T> withGrid(Grid<T> grid) {
      bean.setGrid(grid);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withVisibleProperties(String[] visibleProperties) {
      bean.setVisibleProperties(visibleProperties);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withHeaderConfigs(
        final List<ComponentHeaderConfiguration> headerConfigs) {
      bean.setHeaderConfigs(headerConfigs);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withFooterConfigs(
        final List<ComponentFooterConfiguration> footerConfigs) {
      bean.setFooterConfigs(footerConfigs);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withIntegerFormattingProperties(
        List<String> integerFormattingProperties) {
      bean.setIntegerFormattingProperties(integerFormattingProperties);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withFloatFormattingProperties(List<String> floatFormattingProperties) {
      bean.setFloatFormattingProperties(floatFormattingProperties);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withBooleanFormattingProperties(
        final List<String> booleanFormattingProperties) {
      bean.setBooleanFormattingProperties(booleanFormattingProperties);
      return this;
    }

    ExportExcelComponentConfigurationBuilder<T> withColumnFormatters(
        final Map<Object, ColumnFormatter> columnFormatters) {
      bean.setColumnFormatters(columnFormatters);
      return this;
    }

    ExportExcelComponentConfiguration<T> build() {
      return bean;
    }
  };

  class ExportExcelConfigurationBuilder<T> {
    private ExportExcelConfiguration<T> bean = new ExportExcelConfiguration<T>();

    ExportExcelConfigurationBuilder<T> withGeneratedBy(String generatedBy) {
      bean.setGeneratedBy(generatedBy);
      return this;
    }

    ExportExcelConfigurationBuilder<T> withSheetConfigs(final List<ExportExcelSheetConfiguration<T>> sheetConfigs) {
      bean.setSheetConfigs(sheetConfigs);
      return this;
    }

    ExportExcelConfiguration<T> build() {
      return bean;
    }

  };

  class ExportExcelSheetConfigurationBuilder<T> {

    private ExportExcelSheetConfiguration<T> bean = new ExportExcelSheetConfiguration<T>();

    ExportExcelSheetConfigurationBuilder<T> withReportTitle(String reportTitle) {
      bean.setReportTitle(reportTitle);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withSheetName(String sheetname) {
      bean.setSheetName(sheetname);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withComponentConfigs(
        final List<ExportExcelComponentConfiguration<T>> componentConfigs) {
      bean.setComponentConfigs(componentConfigs);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withIsHeaderSectionRequired(Boolean isHeaderSectionRequired) {
      bean.setIsHeaderSectionRequired(isHeaderSectionRequired);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withIsDefaultSheetTitleRequired(Boolean isDefaultSheetTitleRequired) {
      bean.setIsDefaultSheetTitleRequired(isDefaultSheetTitleRequired);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withDateFormat(String dateFormat) {
      bean.setDateFormat(dateFormat);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withFrozenColumns(int frozenColumns) {
      bean.setFrozenColumns(frozenColumns);
      return this;
    }

    ExportExcelSheetConfigurationBuilder<T> withFrozenRows(int frozenRows) {
      bean.setFrozenRows(frozenRows);
      return this;
    }

    ExportExcelSheetConfiguration<T> build() {
      return bean;
    }
  };

  class MergedCellBuilder {
    private MergedCell bean = new MergedCell();

    MergedCellBuilder withStartProperty(String startProperty) {
      bean.setStartProperty(startProperty);
      return this;
    }

    MergedCellBuilder withEndProperty(String endProperty) {
      bean.setEndProperty(endProperty);
      return this;
    }

    MergedCellBuilder withHeaderKey(String headerKey) {
      bean.setHeaderKey(headerKey);
      return this;
    }

    MergedCell build() {
      return bean;
    }
  };
}
