package org.vaadin.addons.excelexporter.function;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * This function helps you to create the cellstyle you want for a specific cell
 * via columndId, value and rowNum.
 */
@FunctionalInterface
public interface DataCellStyleGeneratorFunction {

	/**
	 * Applies this function to the given arguments.
	 *
	 * @param workbook
	 *            the workbook
	 * @param columnId
	 *            the columnId
	 * @param value
	 *            the value
	 * @param rowNum
	 *            the rowNum
	 * @return the function result
	 */
	CellStyle apply(Workbook workbook, String columnId, Object value, int rowNum);
}