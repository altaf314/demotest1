package com.accolite.gssservice.service;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;

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

		updateWorkbook(workbook, differences);
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
					if(Objects.nonNull(dataToUpdate)) {
						log.info("index: {}-{}, data: {}", i,j,dataToUpdate);
						cell.setCellValue(dataToUpdate);
					}
				}
			}
		}

	}

	public static void compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2, Map<String, String> differences) throws Exception {
		for (int i = 0; i <= sheet1.getPhysicalNumberOfRows(); i++) {
			XSSFRow row = sheet1.getRow(i);
			if (Objects.nonNull(row)) {
				int columns = row.getLastCellNum();
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
		
		
//		int rowNumber = 0;
//		while (sheet1Rows.hasNext()) {
//			Row sheet1Row = sheet1Rows.next();
//			Row sheet2Row = sheet2Rows.next();
//			
//			if (rowNumber == 0) {
//				rowNumber++;
//				continue;
//			}
//			
//	        Iterator<Cell> sheet1RowCells = sheet1Row.iterator();
//
//	        int cellNumber = 0;
//	        while (sheet1RowCells.hasNext()) {
//	        	
//	        }
//	        
//	        rowNumber++;
//		}

//			if ((sheet1row == null && sheet2row == null) && (sheet1row == null || sheet2row == null)) {
//				return;
//			}
//
//			int firstsheet1cell = 0;
//			int lastsheet1cell = sheet1row.getLastCellNum();

//			for (int j = firstsheet1cell; j <= lastsheet1cell; j++) {
//				System.out.println(i);
//				System.out.println(j);
//				log.info(""+i);
//				log.info(""+j);
//				XSSFCell sheet1cell = sheet1row.getCell(j);
//				XSSFCell sheet2cell = sheet2row.getCell(j);
//
//				if ((sheet1cell == null && sheet2cell == null) && (sheet1cell == null || sheet2cell == null)) {
//					return;
//				}
//
//
//				CellType type1 = sheet1cell.getCellType();
//				CellType type2 = sheet2cell.getCellType();

//				if (type1 == type2) {
//					if (CellType.NUMERIC.equals(sheet1cell.getCellType())) {
//						log.info(i + "-" + j + " :: " + sheet1cell.getNumericCellValue() + " -> "
//								+ sheet2cell.getNumericCellValue());
//						if (sheet1cell.getNumericCellValue() != sheet2cell.getNumericCellValue()) {
//							differences.put(i + "-" + j,
//									sheet1cell.getNumericCellValue() + " -> " + sheet2cell.getNumericCellValue());
//						}
//					}
//					if (CellType.STRING.equals(sheet1cell.getCellType())) {
//						log.info(i + "-" + j + " :: " + sheet1cell.getStringCellValue() + " -> "
//								+ sheet2cell.getStringCellValue());
//						if (!(sheet1cell.getStringCellValue().equals(sheet2cell.getStringCellValue()))) {
//							differences.put(i + "-" + j,
//									sheet1cell.getStringCellValue() + " -> " + sheet2cell.getStringCellValue());
//						}
//					}
//					if (CellType.BOOLEAN.equals(sheet1cell.getCellType())) {
//						log.info(i + "-" + j + " :: " + sheet1cell.getBooleanCellValue() + " -> "
//								+ sheet2cell.getBooleanCellValue());
//						if (sheet1cell.getBooleanCellValue() != sheet2cell.getBooleanCellValue()) {
//							differences.put(i + "-" + j,
//									sheet1cell.getBooleanCellValue() + " -> " + sheet2cell.getBooleanCellValue());
//						}
//					}
//				}
//			}
			
//		}
	}
	
	public static String getValue(int x, int y, XSSFSheet sheet){
	    Row row = sheet.getRow(y);
	    if(row==null) return "";
	    Cell cell = row.getCell(x);
	    if(cell==null) return "";
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
		
		
//		for (Row row : sheet) {
//			for (Cell cell : row) {
//				switch (cell.getCellType()) {
//				case STRING:
//					log.info("" + cell.getStringCellValue());
//					break;
//				case NUMERIC:
//					log.info("" + cell.getNumericCellValue());
//					break;
//				case BOOLEAN:
//					log.info("" + cell.getBooleanCellValue());
//					break;
//				case FORMULA:
//					log.info("" + cell.getCellFormula());
//					break;
//				default:
//					log.info("");
//				}
//			}
//		}
	}
	

}
