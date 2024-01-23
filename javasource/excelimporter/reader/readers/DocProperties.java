package excelimporter.reader.readers;

import excelimporter.proxies.DataSource;
import excelimporter.proxies.MappingType;

public class DocProperties {
	private final DataSource dataSource;
	private final MappingType mappingType;
	private final String columnAlias;
	private String staticStringValue = null;

	public DocProperties(DataSource dataSource, MappingType mappingType, String columnAlias) {
		this.dataSource = dataSource;
		this.mappingType = mappingType;
		this.columnAlias = columnAlias;
	}

	public void setStaticStringValue(String value) {
		staticStringValue = value;
	}
	public String getStaticStringValue() {
		return staticStringValue;
	}

	public String getColumnAlias() {
		return columnAlias;
	}

	public DataSource getDataSource() {
		return dataSource;
	}

	public MappingType getMappingType() {
		return mappingType;
	}
}
