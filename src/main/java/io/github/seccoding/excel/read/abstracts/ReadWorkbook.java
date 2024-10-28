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

public abstract class ReadWorkbook<T> extends Readable<T> {

	private String sheetName;
	private int startRow;
	private int physicalNumberOfRows;
	private int physicalNumberOfCells;

	protected ReadWorkbook(Class<T> resultClass) {
		super(resultClass);
		extractSheetName();
		extractStartRow();
	}

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
			this.sheet = getSheet();
			this.physicalNumberOfRows = this.sheet.getPhysicalNumberOfRows();
			return this.workbook;
		}

		throw new RuntimeException(workbookPath.toString() + " isn't excel file format");
	}

	private void extractSheetName() {
		if (this.resultClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheetAnnotation = this.resultClass.getAnnotation(ExcelSheet.class);
			this.sheetName = excelSheetAnnotation.value();
		}
	}

	private void extractStartRow() {
		if (this.resultClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet excelSheetAnnotation = this.resultClass.getAnnotation(ExcelSheet.class);
			this.startRow = excelSheetAnnotation.startRow();
		}
	}

	private Sheet getSheet() {
		if (this.sheetName == null || this.sheetName.length() == 0) {
			return this.workbook.getSheetAt(0);
		}
		return this.workbook.getSheet(this.sheetName);
	}

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

	protected List<T> setValueInExcel() {
		List<T> list = new ArrayList<>();
		Map<String, String> fieldAnnotations = extractFieldAnnotation();

		for (int i = this.startRow; i < this.physicalNumberOfRows; i++) {
			Row row = super.sheet.getRow(i);
			this.physicalNumberOfCells = row.getPhysicalNumberOfCells();

			T newInstance = InstanceUtil.createInstance(resultClass);
			list.add(newInstance);

			for (int c = 0; c < this.physicalNumberOfCells; c++) {
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
					
					arg = convertCellValueToFieldType(arg, field.getType());

					InstanceUtil.invokeMethod(newInstance, setter, arg);
				}
			}
		}

		return list;
	}

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
