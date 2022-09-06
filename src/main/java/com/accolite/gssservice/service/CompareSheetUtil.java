package com.accolite.gssservice.service;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;

import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CompareSheetUtil {

	public static Workbook compareFile(MultipartFile file1, MultipartFile file2) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook(file1.getInputStream());
		XSSFWorkbook workbook2 = new XSSFWorkbook(file2.getInputStream());

		Map<String, String> differences = new HashMap<>();
		if (Objects.nonNull(workbook) && Objects.nonNull(workbook2)) {
			compareTwoSheets(workbook.getSheetAt(0), workbook2.getSheetAt(0), differences);
		}

		if (MapUtils.isNotEmpty(differences)) {
			updateWorkbook(workbook, differences);
		} else {
			throw new Exception("No difference Found !");
		}
		return workbook;
	}

	private static void updateWorkbook(XSSFWorkbook workbook, Map<String, String> differences) {
		XSSFSheet sheet = workbook.getSheetAt(0);

		for (int i = 0; i <= sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet.getRow(i);
			if (Objects.nonNull(row)) {
				int columns = row.getLastCellNum();
				for (int j = 0; j <= columns; j++) {
					Cell cell = row.getCell(j);
					String dataToUpdate = differences
							.get(new StringBuilder().append(j).append(":").append(i).toString());
					if (Objects.nonNull(dataToUpdate)) {
						log.info("index: {}-{}, data: {}", i, j, dataToUpdate);
						cell.setCellValue(dataToUpdate);
					}
				}
			}
		}
	}

	public static void compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2, Map<String, String> differences)
			throws Exception {
		int rows = sheet1.getPhysicalNumberOfRows();
		int columns = sheet1.getRow(0).getLastCellNum();
		if (sheet2.getPhysicalNumberOfRows() > rows) {
			rows = sheet2.getPhysicalNumberOfRows();
		}
		if (sheet2.getRow(0).getLastCellNum() > columns) {
			columns = sheet2.getRow(0).getLastCellNum();
		}

		for (int i = 0; i <= rows; i++) {
			XSSFRow row = sheet1.getRow(i);
			if (Objects.nonNull(row)) {
				for (int j = 0; j <= columns; j++) {
					String data = getValue(i, j, sheet1);
					String data2 = getValue(i, j, sheet2);
					if (!data.equals(data2)) {
						log.info("DIFFERENCE");
						differences.put(new StringBuilder().append(i).append(":").append(j).toString(),
								new StringBuilder().append(data).append(" -> ").append(data2).toString());
					}
				}
			}
		}
	}

	public static String getValue(int x, int y, XSSFSheet sheet) {
		Row row = sheet.getRow(y);
		if (row == null)
			return "";
		Cell cell = row.getCell(x);
		if (cell == null)
			return "";
		return getCellValue(cell);
	}

	public static String getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			return cell.getNumericCellValue() + "";
		case STRING:
			return cell.getStringCellValue();
		case FORMULA:
			return cell.getCellFormula();
		case BLANK:
			return "";
		case BOOLEAN:
			return cell.getBooleanCellValue() + "";
		case ERROR:
			return cell.getErrorCellValue() + "";
		default:
			return "";
		}
	}

	public static void readSheet(Sheet sheet) {
		Iterator<Row> rows = sheet.iterator();

		while (rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cells = row.iterator();

			while (cells.hasNext()) {
				Cell cell = cells.next();
				log.info(getCellValue(cell));
			}

		}
	}

}
