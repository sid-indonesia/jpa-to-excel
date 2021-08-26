package org.sidindonesia.jpatoexcel.helper;

import java.beans.PropertyDescriptor;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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

			fetchAllRowsFromDatabaseAndFillToExcel(context, entityClass, sheet, fields, getterMethods,
				jpaRepositoryPackageName);
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

	private static void fetchAllRowsFromDatabaseAndFillToExcel(ApplicationContext context, Class<?> entityClass,
		Sheet sheet, Field[] fields, Map<String, Method> getterMethods, String jpaRepositoryPackageName) {
		Class<?> repositoryClass = ReflectionsUtil.getRepositoryClassOfEntity(entityClass, jpaRepositoryPackageName);
		try {
			Object repositoryInstance = context.getBean(repositoryClass);
			Object invokeResult = repositoryClass.getMethod("findAll").invoke(repositoryInstance);
			List<?> result = (List<?>) invokeResult;

			AtomicInteger rowIdx = new AtomicInteger();
			result.stream().forEach(entry -> fillContentRows(sheet, fields, rowIdx, entry, getterMethods));
		} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException | NoSuchMethodException
			| SecurityException e) {
			throw new ExcelWriteException("Failed to invoke method `findAll` of class: " + repositoryClass + "\nCause: "
				+ e.getCause() + "\nMessage: " + e.getMessage());
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
}