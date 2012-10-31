package com.cmhk.export

import groovy.lang.Closure;
import groovy.text.SimpleTemplateEngine

import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
/**
 * Excel导出builder
 * @author Tong
 *
 */
class ExcelBuilder {

	Workbook workbook
	private Sheet sheet
	private Row row
	private Cell cell

	def engine = new SimpleTemplateEngine()


	public ExcelBuilder(String path) {
		new File(path).withInputStream{is->
			workbook = WorkbookFactory.create(is)
		}
	}

	public ExcelBuilder(File template){
		template.withInputStream{is->
			workbook = WorkbookFactory.create(is)
		}
		
	}

	public ExcelBuilder(byte[] template){
		workbook = WorkbookFactory.create(new ByteArrayInputStream(template))
	}


	/**
	 * Creates a new workbook.
	 *
	 * @param the closure holds nested {@link ExcelFile} method calls
	 * @return the created {@link Workbook}
	 */
	Workbook workbook(Closure closure) {
		assert closure

		closure.delegate = this
		closure.call()
		workbook
	}


	void sheet(String name, Closure closure) {
		assert workbook

		assert name
		assert closure

		sheet = workbook.getSheet(name)
		if(!sheet)
			sheet = workbook.createSheet(name)

		closure.delegate = sheet
		closure.call()
	}
	
	void sheet(String name,String newName,Closure closure){
		assert workbook
		assert name
		assert closure

		sheet = workbook.getSheet(name)
		if(!sheet)
			sheet = workbook.createSheet(name)
		def sheetIndex = workbook.getSheetIndex(sheet)
		workbook.setSheetName(sheetIndex,newName)
		closure.delegate = sheet
		closure.call()
	}
	
	void cSheet(String name,String newName,Closure closure){
		assert workbook
		assert name
		assert closure
		
		def sheetIndex = workbook.getSheetIndex( workbook.getSheet(name))
		sheet = workbook.cloneSheet(sheetIndex)
		workbook.setSheetName(workbook.getSheetIndex(sheet),newName)
		closure.delegate = sheet
		closure.call()
	}

	void row(int rowIndex,Closure closure){
		assert sheet
		assert rowIndex > 0
		assert closure

		row = sheet.getRow(rowIndex-1)
		if(!row){
			row = sheet.createRow(rowIndex-1)
		}
		closure.delegate = row
		closure.call()
	}
	
	void iRow(int rowIndex,Closure closure){
		assert sheet
		assert rowIndex > 0
		assert closure

		insertRow(sheet,rowIndex-1, 1)
		row = sheet.getRow(rowIndex-1)
		if(!row){
			row = sheet.createRow(rowIndex-1)
		}
		closure.delegate = row
		closure.call()
	}

	void cell(String colString,value){
		assert row
		assert colString ==~/[A-Z]/

		cell = row.getCell(CellReference.convertColStringToIndex(colString))
		setCell(cell, value)
	}

	void cell(String colString,Closure closure){
		assert row
		assert colString ==~/[A-Z]/
		assert closure

		cell = row.getCell(CellReference.convertColStringToIndex(colString))
		closure.delegate = cell
		closure.call()
	}

	void template(binding){
		assert cell
		def value = engine.createTemplate(getCellValue(cell).toString()).make(binding).toString()
		setCell(cell, value)
	}




	void setCell(Cell cell,value){
		switch (value) {
			case Date: cell.setCellValue((Date) value); break
			case Double: cell.setCellValue((Double) value); break
			case BigDecimal: cell.setCellValue(((BigDecimal) value).doubleValue()); break
			case Number: cell.setCellValue(((Number) value).doubleValue()); break
			default:
				def stringValue = value?.toString() ?: ""
				if (stringValue.startsWith('=')) {
					cell.setCellType(Cell.CELL_TYPE_FORMULA)
					cell.setCellFormula(stringValue.substring(1))
				} else {
					cell.setCellValue(new HSSFRichTextString(stringValue))
				}
				break
		}
	}

	def getCellValue(Cell cell) {

		def cellValue
		switch (cell.cellType) {
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.stringCellValue
				break;
			case Cell.CELL_TYPE_NUMERIC:

				if (DateUtil.isCellDateFormatted(cell)) {
					cellValue= cell.dateCellValue
				} else {

					cellValue= cell.numericCellValue
				}
				break;
			case Cell.CELL_TYPE_ERROR:
				return null
			case Cell.CELL_TYPE_FORMULA:
				return null
			case Cell.CELL_TYPE_BOOLEAN:
				return cell.booleanCellValue
			case Cell.CELL_TYPE_BLANK:
				return null;
			default:break;
		}
		return cellValue;
	}

	void insertRow(Sheet sheet, int starRow, int rows) {
		sheet.shiftRows(starRow + 1, sheet.getLastRowNum(), rows, true, false);
		starRow = starRow - 1;
		for (int i = 0; i < rows; i++) {
			Row sourceRow = null;
			Row targetRow = null;
			Cell sourceCell = null;
			Cell targetCell = null;
			short m;
			starRow = starRow + 1;
			sourceRow = sheet.getRow(starRow);
			targetRow = sheet.createRow(starRow + 1);
			targetRow.setHeight(sourceRow.getHeight());
			for (m = sourceRow.getFirstCellNum(); m < sourceRow.getLastCellNum(); m++) {
				sourceCell = sourceRow.getCell(m);
				targetCell = targetRow.createCell(m);
				//targetCell.setEncoding(sourceCell.getEncoding());
				targetCell.setCellStyle(sourceCell.getCellStyle());
				targetCell.setCellType(sourceCell.getCellType());
			}
		}
	}
}
