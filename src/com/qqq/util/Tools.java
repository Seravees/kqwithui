package com.qqq.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qqq.model.Var;

public class Tools {
	public static Workbook open(String path) throws IOException {
		File file = new File(path);

		InputStream stream = null;

		stream = new FileInputStream(file);

		Workbook wb = null;
		if (path.substring(path.lastIndexOf(".") + 1).equals("xls")) {
			wb = new HSSFWorkbook(stream);
		} else if (path.substring(path.lastIndexOf(".") + 1).equals("xlsx")) {
			wb = new XSSFWorkbook(stream);
		} else {
		}

		return wb;
	}

	public static Workbook openPerson(String fileName, String fileType)
			throws IOException {
		File file = new File(Var.getOutPath());
		File[] files = file.listFiles();
		InputStream stream = null;

		for (int i = 0; i < files.length; i++) {
			if (files[i].getName().equals(fileName + "." + fileType)) {
				stream = new FileInputStream(files[i]);
			} else if (files[i].getName().contains(
					fileName.split("-")[1] + "." + fileType)) {
				stream = new FileInputStream(files[i]);
			}
		}

		Workbook wb = null;
		if (fileType.equals("xls")) {
			wb = new HSSFWorkbook(stream);
		} else if (fileType.equals("xlsx")) {
			wb = new XSSFWorkbook(stream);
		} else {
		}
		System.out.println("-" + fileName + "-");

		return wb;

	}

	@SuppressWarnings("deprecation")
	public static void writeString(Workbook wb, int rownum, int cellnum,
			String string, boolean flag) {
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		if (cell == null) {
			cell = row.createCell(cellnum);
		}

		CellStyle style = wb.createCellStyle();
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(HSSFColor.YELLOW.index);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);

		CellStyle style2 = wb.createCellStyle();
		style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style2.setFillForegroundColor(HSSFColor.WHITE.index);
		style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2.setWrapText(true);

		CellStyle style3 = wb.createCellStyle();
		style3.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		if (flag) {
			if (string.contains("缺勤") || string.contains("否")) {
				cell.setCellStyle(style);
			} else {
				cell.setCellStyle(style2);
			}
		} else {
			cell.setCellStyle(style3);
		}

		System.out.println(rownum + "," + cellnum + ":" + string);

		cell.setCellValue(string);
	}

	@SuppressWarnings("deprecation")
	public static void writeDouble(Workbook wb, int rownum, int cellnum,
			double string, boolean flag) {
		Sheet sheet = wb.getSheetAt(0);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		if (cell == null) {
			cell = row.createCell(cellnum);
		}

		CellStyle style = wb.createCellStyle();

		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);

		CellStyle style2 = wb.createCellStyle();

		style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		if (flag) {
			cell.setCellStyle(style);
		} else {
			cell.setCellStyle(style2);
		}

		if (cell != null) {
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
		}

		System.out.println(rownum + "," + cellnum + ":" + string);

		cell.setCellValue(string);

	}

	@SuppressWarnings("deprecation")
	public static void shift(Workbook wb, int startRow) {
		Sheet sheet = wb.getSheetAt(0);
		sheet.shiftRows(startRow, sheet.getLastRowNum(), 1, true, false);
		Row row = sheet.createRow(startRow);
		row.setHeight(sheet.getRow(startRow - 1).getHeight());
		for (int i = 0; i <= 4; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(sheet.getRow(startRow - 1).getCell(i)
					.getCellStyle());
			switch (sheet.getRow(startRow - 1).getCell(i).getCellType()) {
			case HSSFCell.CELL_TYPE_NUMERIC:
				cell.setCellValue(sheet.getRow(startRow - 1).getCell(i)
						.getNumericCellValue());
				break;
			case HSSFCell.CELL_TYPE_STRING:
				cell.setCellValue(sheet.getRow(startRow - 1).getCell(i)
						.getStringCellValue());
				break;
			default:
				break;
			}
		}
		for (int i = 5; i <= 14; i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(sheet.getRow(startRow - 1).getCell(i)
					.getCellStyle());
		}
	}

	public static void create(String path, String fileName, String fileType,
			Workbook wb) throws IOException {
		File file1 = new File(path);
		if (!file1.exists()) {
			file1.mkdir();
		}
		File file = new File(path + "/" + fileName + "." + fileType);
		if (!file.exists()) {
			file.createNewFile();
		}
		OutputStream stream = new FileOutputStream(file);
		wb.write(stream);
		stream.close();
	}

	public static void savePerson(String fileName, String fileType, Workbook wb)
			throws IOException {
		File file = new File(Var.getOutPath());
		File[] files = file.listFiles();

		OutputStream stream = null;

		for (int i = 0; i < files.length; i++) {
			if (files[i].getName().equals(fileName + "." + fileType)) {
				stream = new FileOutputStream(files[i]);
			} else if (files[i].getName().contains(
					fileName.split("-")[1] + "." + fileType)) {
				stream = new FileOutputStream(files[i]);
			}
		}

		wb.write(stream);
		stream.close();

	}

	@SuppressWarnings({ "deprecation", "resource" })
	public static List<List<Object>> readAll(String path) throws IOException {
		List<List<Object>> objects = new ArrayList<List<Object>>();
		if (!path.equals("")) {
			File file = new File(path);
			if (file.exists()) {
				InputStream stream = new FileInputStream(file);
				// InputStream stream = new FileInputStream(path + fileName +
				// "."
				// + fileType);
				Workbook wb = null;

				if (path.substring(path.lastIndexOf(".") + 1).equals("xls")) {
					wb = new HSSFWorkbook(stream);
				} else if (path.substring(path.lastIndexOf(".") + 1).equals(
						"xlsx")) {
					wb = new XSSFWorkbook(stream);
				} else {
				}
				Sheet sheet = wb.getSheetAt(0);
				// System.out.println(sheet.getLastRowNum());
				for (Row row : sheet) {
					List<Object> temp = new ArrayList<Object>();
					for (Cell cell : row) {
						switch (cell.getCellType()) {
						case HSSFCell.CELL_TYPE_STRING:
							temp.add(cell.getStringCellValue());
							// System.out.println(cell.getStringCellValue());
							break;
						case HSSFCell.CELL_TYPE_NUMERIC:
							temp.add(cell.getNumericCellValue());
							break;
						case HSSFCell.CELL_TYPE_BLANK:
							temp.add(" ");
						default:
							break;
						}
					}
					objects.add(temp);
				}
				return objects;
			} else {
				return null;
			}
		} else {
			return null;
		}
	}

	public static boolean isMerged(Sheet sheet, int row, int column) {
		boolean flag = false;
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					flag = true;
				}
			}
		}
		return flag;
	}

}
