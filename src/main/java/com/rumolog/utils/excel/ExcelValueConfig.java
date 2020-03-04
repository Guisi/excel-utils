package com.rumolog.utils.excel;

import org.apache.commons.lang3.StringUtils;

public class ExcelValueConfig {

	private final String title;
	private final String propertyName;
	private final boolean bold;
	private final boolean capitalize;

	public ExcelValueConfig(String title, String propertyName, boolean bold, boolean capitalize) {
		super();
		if (capitalize) {
			title = StringUtils.capitalize(title);
		}
		this.title = title;
		this.propertyName = propertyName;
		this.bold = bold;
		this.capitalize = capitalize;
	}

	public ExcelValueConfig(String title, String propertyName) {
		this(title, propertyName, false, false);
	}

	public String getTitle() {
		return title;
	}

	public String getPropertyName() {
		return propertyName;
	}

	public boolean isBold() {
		return bold;
	}

	public boolean isCapitalize() {
		return capitalize;
	}

}