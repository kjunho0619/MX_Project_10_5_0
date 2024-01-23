package excelimporter.reader.readers;

import excelimporter.reader.readers.ExcelRowProcessor.ExcelCellData;
import com.mendix.replication.MendixReplicationException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Spliterator;
import java.util.function.Predicate;
import java.util.stream.StreamSupport;

public class ExcelDataReader {

	private static final DataFormatter formatter = new DataFormatter();

	public static long readData(String excelFile, int sheetIndex, int startRowIndex, ExcelRowProcessor rowProcessor, Predicate<String> isColumnUsed)
			throws IOException, SAXException, ExcelRuntimeException {

		try (Workbook workbook = WorkbookFactory.create(new File(excelFile))) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			StreamSupport.stream(sheet.spliterator(), false).forEach(row -> {
				try {
					if (row.getRowNum() >= startRowIndex) {
						rowProcessor.processValues(readRow(row.spliterator(), isColumnUsed), row.getRowNum(), sheetIndex);
					}
				} catch (MendixReplicationException e) {
					throw new ExcelRuntimeException("Unable to store Excel row #" + (row.getRowNum() + 1) + " @Sheet #" + sheetIndex, e);
				}
			});
			return rowProcessor.getRowCounter();
		} catch (ExcelRuntimeException e) {
			throw new SAXException(e.getMessage());
		} finally {
			try {
				rowProcessor.finish();
				if (rowProcessor.getRowCounter() == 0)
					ExcelReader.logNode.warn("Excel Importer could not import any rows. Please check if the template is configured correctly. If the file was not created with Microsoft Excel for desktop, try opening the file with Excel and saving it with the same name before importing.");
				else
					ExcelReader.logNode.info("Excel Importer successfully imported " + rowProcessor.getRowCounter() + " rows");
			} catch (MendixReplicationException e) {
				throw new ExcelRuntimeException(e); // needed for backward compatibility
			}
		}
	}

	private static ExcelCellData[] readRow(Spliterator<Cell> cellIterator, Predicate<String> isColumnUsed) {
		final ArrayList<ExcelCellData> data = new ArrayList<>();
	    StreamSupport.stream(cellIterator, false).forEach(cell -> {
	    	// add skipped columns
	    	final int columnIndex = cell.getColumnIndex();
	    	while (columnIndex > data.size()) {
	    		data.add(null);
			}

			// add column
			final Object rawData = getValue(cell, cell.getCellType());
			final ExcelCellData cellData = ((rawData != null) && isColumnUsed.test(String.valueOf(columnIndex)))
				? evaluateCellData(cell, rawData.toString())
				: null;
			data.add(cellData);
		});
		return data.toArray(new ExcelRowProcessor.ExcelCellData[0]);
	}

	private static Object getValue(Cell cell, CellType cellType) {
		switch (cellType) {
			case STRING:
				return cell.getStringCellValue();
			case NUMERIC:
				return cell.getNumericCellValue();
			case BOOLEAN:
				return cell.getBooleanCellValue() ? "1" : "0";
			case FORMULA:
				return (cell.getCachedFormulaResultType() != null)
					? getValue(cell, cell.getCachedFormulaResultType())
					: cell.getCellFormula();
			case ERROR:
				return cell.getErrorCellValue();
			case BLANK:
			case _NONE:
			default:
				return null;
		}
	}

	private static ExcelCellData evaluateCellData(Cell cell, String cellValueString) {
		if ( ExcelReader.logNode.isTraceEnabled() )
			ExcelReader.logNode.trace("Reading " + cell.getAddress() + " / '" + cellValueString + "' / " + cell.getCellType());

		final int columnIndex = cell.getColumnIndex();
		switch (cell.getCellType()) {
			case BOOLEAN:
				return new ExcelCellData(columnIndex, cellValueString, Integer.parseInt(cellValueString) == 1);
			case ERROR:
				// imported as null, because this can be handled in Mendix
				return new ExcelCellData(columnIndex, cellValueString, "ERROR:" + cellValueString);
			case FORMULA:
				// We have the formula available, but it makes sense not to use it - this.formula.toString());
				return new ExcelCellData(columnIndex, cellValueString, cellValueString);
			case STRING: // We haven't seen this yet.
				XSSFRichTextString rtsi = new XSSFRichTextString(cellValueString);
				return new ExcelCellData(columnIndex, cellValueString, rtsi.toString());
			case NUMERIC:
				final String formatString = cell.getCellStyle().getDataFormatString();
				if (formatString != null) {
					final double dblCellValue = Double.parseDouble(cellValueString);
					final String formattedValue = formatter.formatRawCellContents(dblCellValue, cell.getCellStyle().getDataFormat(), formatString);

					if (ExcelReader.logNode.isTraceEnabled())
						ExcelReader.logNode.trace("Formatting " + cell.getAddress() + " / '" + cellValueString
								+ "' using format: '" + formatString + "' as " + formattedValue);

					return new ExcelCellData(columnIndex, dblCellValue, formattedValue, formatString);
				} else {
					return new ExcelCellData(columnIndex, cellValueString, null);
				}
			default:
				return null;
		}
	}
}
