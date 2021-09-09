package org.sidindonesia.jpatoexcel.helper;

import java.beans.PropertyDescriptor;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.sidindonesia.jpatoexcel.exception.ExcelWriteException;
import org.sidindonesia.jpatoexcel.util.CamelCaseUtil;
import org.sidindonesia.jpatoexcel.util.ReflectionsUtil;
import org.springframework.beans.BeanUtils;
import org.springframework.context.ApplicationContext;

public final class ExcelHelper {

	private ExcelHelper() {
		// Helper class
	}

	private static final String FAILED_TO_IMPORT_DATA_TO_EXCEL_FILE = "Failed to import data to Excel file: ";

	public static ByteArrayInputStream allEntitiesToExcelSheets(ApplicationContext context, String jpaEntityPackageName,
		String jpaRepositoryPackageName) {

		try (Workbook workbook = new SXSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {

			processEachEntityAsAnExcelSheet(context, workbook, jpaEntityPackageName, jpaRepositoryPackageName);

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		} catch (IOException e) {
			throw new ExcelWriteException(FAILED_TO_IMPORT_DATA_TO_EXCEL_FILE + e.getMessage());
		}
	}

	private static void processEachEntityAsAnExcelSheet(ApplicationContext context, Workbook workbook,
		String jpaEntityPackageName, String jpaRepositoryPackageName) {
		Set<Class<?>> entityClasses = ReflectionsUtil.getAllEntityClasses(jpaEntityPackageName);

		entityClasses.forEach(entityClass -> {
			Sheet sheet = workbook.createSheet(CamelCaseUtil.camelToSnake(entityClass.getSimpleName()));

			Field[] fields = entityClass.getDeclaredFields();
			Map<String, Method> getterMethods = getGetterMethodsByFieldName(fields, entityClass);
			createHeaderRow(sheet, fields);

			List<?> result = retrieveAllRowsFromDatabase(context, entityClass, jpaRepositoryPackageName, "findAll");

			AtomicInteger rowIdx = new AtomicInteger();
			result.stream().forEach(entry -> fillContentRows(sheet, fields, rowIdx, entry, getterMethods));
		});
	}

	private static Map<String, Method> getGetterMethodsByFieldName(Field[] fields, Class<?> entityClass) {
		Map<String, Method> getterMethodsByFieldName = new HashMap<>();
		for (Field field : fields) {
			PropertyDescriptor propertyDescriptor = BeanUtils.getPropertyDescriptor(entityClass, field.getName());
			if (propertyDescriptor != null) {
				getterMethodsByFieldName.put(field.getName(), propertyDescriptor.getReadMethod());
			}
		}
		return getterMethodsByFieldName;
	}

	private static void createHeaderRow(Sheet sheet, Field[] fields) {
		Row headerRow = sheet.createRow(0);
		int headerCol = 0;
		for (Field field : fields) {
			Cell cell = headerRow.createCell(headerCol++);
			cell.setCellValue(CamelCaseUtil.camelToSnake(field.getName()));
		}
		sheet.createFreezePane(0, 1);
	}

	private static List<?> retrieveAllRowsFromDatabase(ApplicationContext context, Class<?> entityClass,
		String jpaRepositoryPackageName, String methodName, Object... args) {
		Class<?> repositoryClass = ReflectionsUtil.getRepositoryClassOfEntity(entityClass, jpaRepositoryPackageName);
		try {
			Object repositoryInstance = context.getBean(repositoryClass);
			if (args.length > 0) {
				if (args.length == 2) {
					return (List<?>) repositoryClass.getMethod(methodName, args[0].getClass(), args[1].getClass())
						.invoke(repositoryInstance, args);
				}
				return (List<?>) repositoryClass.getMethod(methodName, args[0].getClass()).invoke(repositoryInstance,
					args);
			} else {
				return (List<?>) repositoryClass.getMethod(methodName).invoke(repositoryInstance);
			}
		} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException | NoSuchMethodException
			| SecurityException e) {
			throw new ExcelWriteException(
				"Failed to invoke method `" + methodName + "` of class: " + repositoryClass + "\nException: ");
		}
	}

	private static void fillContentRows(Sheet sheet, Field[] fields, AtomicInteger rowIdx, Object entry,
		Map<String, Method> getterMethods) {
		Row contentRow = sheet.createRow(rowIdx.incrementAndGet());
		int contentCol = 0;
		for (Field field : fields) {
			Cell cell = contentRow.createCell(contentCol++);
			try {
				Object invokeGetterResult = getterMethods.get(field.getName()).invoke(entry);
				cell.setCellValue(invokeGetterResult == null ? null : invokeGetterResult.toString());
			} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
				throw new ExcelWriteException(FAILED_TO_IMPORT_DATA_TO_EXCEL_FILE + e.getCause());
			}
		}
	}

	public static ByteArrayInputStream validateAllColumnsAreNotEmpty(ApplicationContext context,
		String jpaEntityPackageName, String jpaRepositoryPackageName, LocalDateTime fromDate, LocalDateTime untilDate) {

		try (Workbook workbook = new SXSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {

			Set<Class<?>> entityClasses = ReflectionsUtil.getAllEntityClasses(jpaEntityPackageName);

			entityClasses.forEach(entityClass -> {
				Sheet sheet = workbook.createSheet(CamelCaseUtil.camelToSnake(entityClass.getSimpleName()));

				Field[] fields = entityClass.getDeclaredFields();
				Map<String, Method> getterMethods = getGetterMethodsByFieldName(fields, entityClass);
				createHeaderRow(sheet, fields);
				createLastHeaderRowsForValidationReportColumns(sheet, fields.length);

				List<?> result = retrieveAllRowsFromDatabase(context, entityClass, jpaRepositoryPackageName,
					"findAllByDateCreatedBetween", fromDate, untilDate);

				AtomicInteger rowIdx = new AtomicInteger();
				CellStyle missingValueCellStyle = createMissingValueCellStyle(workbook);
				result.stream().forEach(entry -> fillAndValidateContentRows(sheet, fields, rowIdx, entry, getterMethods,
					missingValueCellStyle));
			});

			workbook.write(out);
			return new ByteArrayInputStream(out.toByteArray());
		} catch (IOException e) {
			throw new ExcelWriteException(FAILED_TO_IMPORT_DATA_TO_EXCEL_FILE + e.getMessage());
		}
	}

	private static CellStyle createMissingValueCellStyle(Workbook workbook) {
		XSSFCellStyle missingValueCellStyle = (XSSFCellStyle) workbook.createCellStyle();
		missingValueCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
		missingValueCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return missingValueCellStyle;
	}

	private static void createLastHeaderRowsForValidationReportColumns(Sheet sheet, int lastHeaderCol) {
		Row headerRow = sheet.getRow(0);
		Cell totalMissingValuesHeaderCell = headerRow.createCell(lastHeaderCol++);
		totalMissingValuesHeaderCell.setCellValue("Total Missing Values");
		Cell filledValuesPercentage = headerRow.createCell(lastHeaderCol++);
		filledValuesPercentage.setCellValue("Percentage of non-missing values"); // fields.length - total missng values
																					// / fields.length * 100
		Cell validationReportCell = headerRow.createCell(lastHeaderCol);
		validationReportCell.setCellValue("Validation Report"); // which column value is missing
	}

	private static void fillAndValidateContentRows(Sheet sheet, Field[] fields, AtomicInteger rowIdx, Object entry,
		Map<String, Method> getterMethods, CellStyle missingValueCellStyle) {
		Row contentRow = sheet.createRow(rowIdx.incrementAndGet());
		int contentCol = 0;
		for (Field field : fields) {
			Cell cell = contentRow.createCell(contentCol++);
			try {
				Object invokeGetterResult = getterMethods.get(field.getName()).invoke(entry);

				if (invokeGetterResult == null) {
					cell.setCellStyle(missingValueCellStyle);
				} else {
					cell.setCellValue(invokeGetterResult.toString());
				}

			} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
				throw new ExcelWriteException(FAILED_TO_IMPORT_DATA_TO_EXCEL_FILE + e.getCause());
			}
		}
	}
}