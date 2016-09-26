package com.yxt;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
	public static void process(String path) {
		OutputStream os = null;
		Workbook workbook = null;
		try {
			workbook = getWorkbook(path);
	    	Sheet sheet = workbook.getSheetAt(0);
	    	int rowNum = sheet.getPhysicalNumberOfRows();
	    	for(int r = 1; r < rowNum; r++) {
	    		Row row = sheet.getRow(r);
	    		Cell cell = row.getCell(0);
	    		cell.setCellType(Cell.CELL_TYPE_STRING);
	    		String cellValue = cell.getStringCellValue();
	    		if(!cellValue.startsWith("0")) {
	    			int length = cellValue.length();
	    			for(int i = 0; i < 8-length; i++) {
	    				cellValue = "0"+cellValue;
	    			}
	    			cell.setCellValue(cellValue);
	    		}
	    	}
	    	os = new FileOutputStream("yxt.xls");
	    	workbook.write(os);
		}catch(Exception e) {
			JOptionPane.showMessageDialog(null, e.getMessage(), "错误", JOptionPane.ERROR_MESSAGE);
			e.printStackTrace();
		}finally {
			try {
				os.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
    	
	}
	public static Workbook getWorkbook(String path) throws IOException {
		FileInputStream fis = new FileInputStream(path);
		Workbook workbook;
		try {
			workbook = new XSSFWorkbook(fis);
			return workbook;
		}catch(Exception e) {
			 fis = new FileInputStream(path);
			 workbook = new HSSFWorkbook(fis);
			 return workbook;
		}
		
	}
	public static String getCellContent(Sheet sheet, int i, int col) {
		Cell meCell = isMergedRegion(sheet, i, col);
		if(meCell != null) {
			return getCellValue(meCell);
		}
		String cellContent = "";
		Row row = sheet.getRow(i);
		if(row!=null) {
			Cell cell = row.getCell(col);
			cellContent = getCellValue(cell);
			if("同上".equals(cellContent)) {
				return getCellContent(sheet, i-1, col);
			}
		}
		return cellContent;
	}
	/**
	 * 判断单元格是否为合并单元格
	 */
	public static Cell isMergedRegion(Sheet sheet, int i, int col) {
		int n = sheet.getNumMergedRegions();
		for(int k = 0; k < n; k++) {
			CellRangeAddress cra = sheet.getMergedRegion(k);
			int firstCol = cra.getFirstColumn();
			int lastCol = cra.getLastColumn();
			int firstRow = cra.getFirstRow();
			int lastRow = cra.getLastRow();
			if(i >= firstRow && i <= lastRow && col >= firstCol && col <= lastCol) {
				return sheet.getRow(firstRow).getCell(firstCol);
			}
		}
		return null;
	}
	/**
	 * 获取单元格内容
	 */
	public static String getCellValue(Cell cell) {
		String cellContent = "";
		if(cell!=null) {
			int cellType = cell.getCellType();
			switch(cellType) {
			case Cell.CELL_TYPE_FORMULA: cellContent = cell.getCellFormula(); break;
			case Cell.CELL_TYPE_NUMERIC: 
				if(HSSFDateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellContent = sdf.format(date);
				}else if(cell.getCellStyle().getDataFormat()==176||cell.getCellStyle().getDataFormat()==57){
					Date date = cell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM");
					cellContent = sdf.format(date);
				}else {
					cellContent = cell.getNumericCellValue()+"";
					if(cellContent.equals("36647.0")) System.out.println(cell.getCellStyle().getDataFormat());
				}
				break;
			case Cell.CELL_TYPE_STRING: cellContent = cell.getStringCellValue();break;
			default: cellContent = cell.getStringCellValue();break;
			}
		}
		cellContent = cellContent.replace("\t", "");
		cellContent = cellContent.replace("\n", "");
		cellContent = cellContent.replace("　", "");
		cellContent = cellContent.trim();
		return cellContent;
	}
}
