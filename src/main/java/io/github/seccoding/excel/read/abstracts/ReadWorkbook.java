package io.github.seccoding.excel.read.abstracts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.github.seccoding.excel.annotations.ExcelSheet;
import io.github.seccoding.excel.util.FileType;
import io.github.seccoding.excel.util.InstanceUtil;

/**
 * 워크북 시트의 값들을 읽어 냄.
 * @param <T>
 */
public abstract class ReadWorkbook<T> extends Readable<T> {

	/**
	 * 읽을 시트 명
	 */
	private String sheetName;
	
	/**
	 * 읽어낼 시작 행 번호
	 */
	private int startRow;
	
	protected ReadWorkbook(Class<T> resultClass) {
		super(resultClass);
		this.extractSheetName();
		this.extractStartRow();
	}

	/**
	 * 엑셀파일을 읽어 워크북으로 생성. 
	 * @param workbookPath
	 * @return
	 */
	protected Workbook loadWorkbook(Path workbookPath) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(workbookPath.toFile());
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e.getMessage(), e);
		}

		if (FileType.isXls(workbookPath.toString())) {
			try {
				this.workbook = new HSSFWorkbook(fis);
			} catch (IOException e) {
				throw new RuntimeException(e.getMessage(), e);
			} finally {
				if (fis != null) {
					try {
						fis.close();
					} catch (IOException e1) {
					}
				}
			}
		}
		if (FileType.isXlsx(workbookPath.toString())) {
			try {
				this.workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
				throw new RuntimeException(e.getMessage(), e);
			} finally {
				if (fis != null) {
					try {
						fis.close();
					} catch (IOException e1) {
					}
				}
			}
		}

		if (fis != null) {
			try {
				fis.close();
			} catch (IOException e1) {
			}
		}

		if (this.workbook != null) {
			super.sheet = this.getSheet();
			this.sheetName = super.sheet.getSheetName();
			
			super.sheetList = new ArrayList<>();
			int sheetCount = super.workbook.getNumberOfSheets();
			for (int i = 0; i < sheetCount; i++) {
				super.sheetList.add(super.workbook.getSheetAt(i));
			}
			
			return this.workbook;
		}

		throw new RuntimeException(workbookPath.toString() + " isn't excel file format");
	}
	
	/**
	 * @ExcelSheet 에서 읽을 Sheet명을 조회.
	 */
	private void extractSheetName() {
		if (this.resultClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheetAnnotation = this.resultClass.getAnnotation(ExcelSheet.class);
			this.sheetName = excelSheetAnnotation.value();
		}
	}

	/**
	 * 엑셀에 존재하는 모든 시트 목록을 반환.
	 * @return
	 */
	protected List<Sheet> getAllSheets() {
		return super.sheetList;
	}
	
	/**
	 * @ExcelSheet 에서 읽을 시작 행 번호를 조회.
	 */
	private void extractStartRow() {
		if (this.resultClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheetAnnotation = this.resultClass.getAnnotation(ExcelSheet.class);
			this.startRow = excelSheetAnnotation.startRow();
		}
	}

	/**
	 * 워크북에서 읽을 시트 추출.
	 * 시트명이 없을 경우 첫 번째 시트를 추출한다.
	 * @return
	 */
	protected Sheet getSheet() {
		if (this.sheetName == null || this.sheetName.length() == 0) {
			return super.workbook.getSheetAt(0);
		}
		return super.workbook.getSheet(this.sheetName);
	}

	/**
	 * "@Field가 적용된 멤버변수를 탐색"
	 * @return Map<컬럼명, 멤버변수명>
	 */
	private Map<String, String> extractFieldAnnotation() {
		Field[] fields = super.resultClass.getDeclaredFields();
		Map<String, String> fieldNames = new HashMap<>();
		for (Field field : fields) {
			if (super.isPresentFieldAnnotation(field)) {
				io.github.seccoding.excel.annotations.Field fieldAnnotation = super.getFieldAnnotation(field);
				fieldNames.put(fieldAnnotation.value(), field.getName());
			}

		}

		return fieldNames;
	}

	/**
	 * @ExcelSheet 에 정의한 시트명과 행번호 부터 데이터를 읽어온다.
	 * @return 시트의 내용.
	 */
	protected List<T> setValueInExcel() {
		return setValueInExcel(super.sheet, this.startRow);
	}
	
	/**
	 * 엑셀 워크북의 시트를 읽어 리스트로 변환.
	 * @param sheet 읽으려는 시트
	 * @param startRow 시트에서 읽기 시작할 행 번호
	 * @return 시트의 내용.
	 */
	protected List<T> setValueInExcel(Sheet sheet, int startRow) {
		List<T> list = new ArrayList<>();
		Map<String, String> fieldAnnotations = this.extractFieldAnnotation();
		
		for (int i = startRow; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			int physicalNumberOfCells = row.getPhysicalNumberOfCells();

			T newInstance = InstanceUtil.createInstance(super.resultClass);
			list.add(newInstance);

			for (int c = 0; c < physicalNumberOfCells; c++) {
				Cell cell = row.getCell(c);
				String cellName = CellReference.convertNumToColString(cell != null ? cell.getColumnIndex() : c);
				cellName = cellName.replace("[0-9]+", "");

				if (fieldAnnotations.containsKey(cellName)) {
					String fieldName = fieldAnnotations.get(cellName);
					Field field = InstanceUtil.getFieldInNewInstance(newInstance, fieldName);

					Method setter = InstanceUtil.getMethod(newInstance, "set", fieldName, field.getType());

					Object arg = getCellValue(cell);

					io.github.seccoding.excel.annotations.Field fieldAnnotation = field
							.getAnnotation(io.github.seccoding.excel.annotations.Field.class);
					if (fieldAnnotation.isDate() && arg instanceof Date) {
						SimpleDateFormat sdf = new SimpleDateFormat(fieldAnnotation.dateFormat());
						arg = sdf.format(arg);
					}
					
					arg = this.convertCellValueToFieldType(arg, field.getType());

					InstanceUtil.invokeMethod(newInstance, setter, arg);
				}
			}
		}

		return list;
	}

	/**
	 * 컬럼의 타입 별로 값을 추출.
	 * @param cell 추출하려는 셀.
	 * @return
	 */
	private Object getCellValue(Cell cell) {
		if (cell.getCellType() == CellType.FORMULA) {
			return cell.getCellFormula();
		} else if (cell.getCellType() == CellType.BLANK) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == CellType.BOOLEAN) {
			return cell.getBooleanCellValue();
		} else if (cell.getCellType() == CellType.ERROR) {
			return cell.getErrorCellValue();
		} else if (cell.getCellType() == CellType.NUMERIC) {
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		}

		return cell.getStringCellValue();
	}

	private Object convertCellValueToFieldType(Object value, Class<?> classType) {
		if (classType == byte.class) {
			double doubleArg = Double.parseDouble(String.valueOf(value));
			return (byte) doubleArg;
		} else if (classType == short.class) {
			double doubleArg = Double.parseDouble(String.valueOf(value));
			return (short) doubleArg;
		} else if (classType == int.class) {
			double doubleArg = Double.parseDouble(String.valueOf(value));
			return (int) doubleArg;
		} else if (classType == long.class) {
			double doubleArg = Double.parseDouble(String.valueOf(value));
			return (long) doubleArg;
		} else if (classType == float.class) {
			return Double.parseDouble(String.valueOf(value));
		} else if (classType == double.class) {
			return Double.parseDouble(String.valueOf(value));
		} else if (classType == boolean.class) {
			return Boolean.parseBoolean(String.valueOf(value));
		}
		return String.valueOf(value);
	}

}
