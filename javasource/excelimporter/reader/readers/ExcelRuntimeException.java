package excelimporter.reader.readers;

public class ExcelRuntimeException extends RuntimeException {
	private static final long serialVersionUID = 123456L;

	public ExcelRuntimeException(Exception e) {
		super(e);
	}

	public ExcelRuntimeException(String message, Exception e) {
		super(message, e);
	}
}
