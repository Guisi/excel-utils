package com.rumolog.utils.excel;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.time.Instant;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class ExcelUtils {

	private Map<Class<? extends Object>, CellStyle> cellStyleMap = new HashMap<>();

	public <T> ByteArrayOutputStream writeToExcel(List<ExcelValueConfig> configurations, List<T> data) {
		if (configurations.isEmpty() || data.isEmpty()) {
			return null;
		}
		try (HSSFWorkbook workbook = new HSSFWorkbook()) {
			HSSFSheet sheet = workbook.createSheet();

			int rowCount = 0;
			int columnCount = 0;

			// Header
			Row row = sheet.createRow(rowCount++);
			for (ExcelValueConfig config : configurations) {
				Cell cell = row.createCell(columnCount++);
				CellStyle cellStyle = workbook.createCellStyle();

				if (config.isBold()) {
					HSSFFont font = workbook.createFont();
					font.setBold(true);
					cellStyle.setFont(font);
				}

				cell.setCellValue(config.getTitle());
				cell.setCellStyle(cellStyle);
			}

			// Body
			for (T t : data) {
				row = sheet.createRow(rowCount++);
				columnCount = 0;
				for (ExcelValueConfig config : configurations) {
					Cell cell = row.createCell(columnCount);

					Object value = ExcelUtils.getValueFromProperty(t, config.getPropertyName());
					this.setCellValue(workbook, cell, value);
					columnCount++;
				}
			}

			for (int i = 0; i < configurations.size(); i++) {
				sheet.autoSizeColumn(i);
			}

			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			workbook.write(baos);

			return baos;
		} catch (Exception e) {
			throw new IllegalStateException("Erro ao gerar XLS", e);
		}
	}

	private void setCellValue(HSSFWorkbook workbook, Cell cell, Object value) {
		if (value == null) {
			cell.setCellValue("");
		} else {
			if (value instanceof String) {
				cell.setCellValue((String) value);
			} else if (value instanceof Long) {
				cell.setCellValue((Long) value);
			} else if (value instanceof Integer) {
				cell.setCellValue((Integer) value);
			} else if (value instanceof Double) {
				cell.setCellValue((Double) value);
			} else if (value instanceof Instant) {
				Date date = Date.from(((Instant) value));
				cell.setCellValue(date);
				cell.setCellStyle(getCellStyle(Instant.class, workbook));
			} else if (value instanceof BigDecimal) {
				cell.setCellValue(((BigDecimal) value).doubleValue());
			} else if (value instanceof Float) {
				cell.setCellValue((Float) value);
			}
		}
	}

	private CellStyle getCellStyle(Class<? extends Object> type, HSSFWorkbook workbook) {
		this.cellStyleMap.computeIfAbsent(type, key -> {
			CellStyle cellStyle = workbook.createCellStyle();
			cellStyle.setDataFormat((short) 14);
			return cellStyle;
		});
		return this.cellStyleMap.get(type);
	}

	private static Object getValueFromProperty(Object data, String propertyName) {
		Class<? extends Object> classe = data.getClass();
		try {
			Method method = ExcelUtils.getMethod(classe, propertyName);

			if (method != null) {
				return method.invoke(data, (Object[]) null);
			}
		} catch (Exception e) {
			throw new IllegalStateException(e.getMessage(), e);
		}
		return null;
	}

	private static Method getMethod(Class<? extends Object> classe, String propertyName) {
		try {
			return classe.getMethod("get" + StringUtils.capitalize(propertyName));
		} catch (NoSuchMethodException nme) {
			try {
				return classe.getMethod("get" + propertyName);
			} catch (NoSuchMethodException | SecurityException e) {
				return null;
			}
		}
	}
}