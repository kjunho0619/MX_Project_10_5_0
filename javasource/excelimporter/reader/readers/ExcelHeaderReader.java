package excelimporter.reader.readers;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ExcelHeaderReader {

	public static List<ExcelColumn> getHeaderColumns(File excelFile, int sheetIndex, int rowIndex) throws IOException {
		try (Workbook workbook = WorkbookFactory.create(excelFile)) {
			final Sheet sheet = workbook.getSheetAt(sheetIndex);
			final Row headerRow = sheet.getRow(rowIndex);

			return StreamSupport
					.stream(headerRow.spliterator(), false)
					.map(c -> new ExcelColumn(c.getColumnIndex(), c.getStringCellValue()))
					.collect(Collectors.toList());
		}
	}
}
