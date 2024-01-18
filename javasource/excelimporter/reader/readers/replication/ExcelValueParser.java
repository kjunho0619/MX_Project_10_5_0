package excelimporter.reader.readers.replication;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParsePosition;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DateUtil;

import com.mendix.replication.AbstractValueExtractor;
import com.mendix.replication.ParseException;
import com.mendix.systemwideinterfaces.core.meta.IMetaPrimitive.PrimitiveType;

import excelimporter.reader.readers.ExcelRowProcessor.ExcelCellData;
import excelimporter.proxies.constants.Constants;

public class ExcelValueParser extends AbstractValueExtractor {

	private static final HashMap<String, String> displayMaskMap = new HashMap<>();
	static {
		displayMaskMap.put("m/d/yy", "MM/dd/yy");
		displayMaskMap.put("m/d/yy\\ h:mm;@", "MM/dd/yy HH:mm");
		displayMaskMap.put("m/d/yyyy", "MM/dd/yyyy");
		displayMaskMap.put("m/d/yyyy\\ h:mm;@", "MM/dd/yyyy HH:mm");

		displayMaskMap.put("dd\\-mmm\\-yy;@\\", "dd-MMMM-yy");
		displayMaskMap.put("[$-409]dd\\-mmm\\-yy;@\\", "dd-MMMM-yy");
		displayMaskMap.put("dd\\-mmm\\-yyyy;@\\", "dd-MMMM-yyyy");
		displayMaskMap.put("[$-409]dd\\-mmm\\-yyyy;@\\", "dd-MMMM-yyyy");
		displayMaskMap.put("h:mm:ss\\ AM/PM", "hh:mm:ss aa");
		displayMaskMap.put("[$-409]h:mm:ss\\ AM/PM", "hh:mm:ss aa");
		displayMaskMap.put("dddd\\,\\ mmmm\\ dd\\,\\ yyyy", "EEEE, MMMM dd, yyyy");
		displayMaskMap.put("[$-409]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", "EEEE, MMMM dd, yyyy");

		displayMaskMap.put("\"$\"#,##0_);\\(\"$\"#,##0\\)", "#,##0");
		displayMaskMap.put("\"$\"#,##0_);[Red]\\(\"$\"#,##0\\)", "#,##0");

		displayMaskMap.put("\"$\"#,##0.00_);\\(\"$\"#,##0.00\\)", "#,##0.00");
		displayMaskMap.put("\"$\"#,##0.00_);[Red]\\(\"$\"#,##0.00\\)", "#,##0.00");

		displayMaskMap.put("0.0%", "#0.0%");
		displayMaskMap.put("0.00%", "#0.0%");
		displayMaskMap.put("0.000%", "#0.0%");
		displayMaskMap.put("0.0000%", "#0.0%");
	}

	public ExcelValueParser(ExcelReplicationSettings settings) {
		super(settings);
		this.settings = settings;
	}

	private final ExcelReplicationSettings settings;

	@Override
	public ExcelReplicationSettings getSettings() {
		return settings;
	}

	public Object getValueFromDataSet(String column, PrimitiveType type, Object dataSet) throws ParseException {
		final ExcelCellData[] objects = (ExcelCellData[]) dataSet;
		final int columnIndex = Integer.parseInt(column);

		for (ExcelCellData object : objects) {
			if (object != null && object.getColumnIndex() == columnIndex) {
				return getValue(type, column, object);
			}
		}

		return getValue(type, column, null);
	}

	@Override
	public String getKeyValueFromAlias(Object recordDataSet, String keyAlias) throws ParseException {
		return getKeyValueByPrimitiveType(settings.getMemberType(keyAlias), keyAlias,
			getValueFromDataSet(keyAlias, settings.getMemberType(keyAlias), recordDataSet));
	}

	private Object getValue(PrimitiveType type, String column, ExcelCellData cellData) throws ParseException {
		if (cellData == null) {
			return (Constants.getParseEmptyCells() && settings.hasValueParser(column))
				? getValue(type, column, (String) null)
				: null;
		}
		else if (cellData.getFormattedData() != null && cellData.getFormattedData().toString().equals(""))
		{
			return getValue(type, column, cellData.getFormattedData().toString());
		}
		else if (type == PrimitiveType.DateTime) {
			return parseToDateTime(column, cellData);
		}
		else if (type == PrimitiveType.Decimal || type == PrimitiveType.Integer || type == PrimitiveType.Long) {
			final Object parsed = parseToNumber(column, cellData);
			if (!(parsed instanceof BigDecimal))
				throw new ParseException("Could not parse value '" + cellData.getFormattedData() + "' to " + type.name() + " in column #" + (cellData.getColumnIndex() + 1));
			try {
				if (type == PrimitiveType.Long) {
					return ((BigDecimal) parsed).setScale(0, RoundingMode.FLOOR).longValueExact();
				} else if (type == PrimitiveType.Integer) {
					return ((BigDecimal) parsed).setScale(0, RoundingMode.FLOOR).intValueExact();
				} else
					return parsed;
			} catch (ArithmeticException e) {
				throw new ParseException("Error casting " + parsed + " to " + type + ": " + e.getMessage() , e);
			}
		}
		else if (cellData.getFormattedData() != null) {
			try {
				return getValue(type, column, cellData.getFormattedData());
			} catch (Exception ignore) {
				return getValue(type, column, cellData.getRawData());
			}
		}
		else
			return getValue(type, column, cellData.getRawData());
	}

	private Object parseToNumber(String column, ExcelCellData cellData) throws ParseException {
		if (cellData.getRawData() instanceof Number) {
			final boolean isPercentage = cellData.getFormattedData() != null && cellData.getFormattedData().toString().endsWith("%");
			final Number rawData = (Number) cellData.getRawData();
			final BigDecimal value = (isPercentage)
				? BigDecimal.valueOf(rawData.doubleValue()).multiply(BigDecimal.valueOf(100))
				: BigDecimal.valueOf(rawData.doubleValue()).stripTrailingZeros();
			return getValue(PrimitiveType.Decimal, column, value);
		} else {
			final String number = (cellData.getFormattedData() != null) ? cellData.getFormattedData().toString() : cellData.getRawData().toString();
			final ParsePosition position = new ParsePosition(0);
			final DecimalFormat numberFormat = (DecimalFormat) NumberFormat.getInstance(Locale.US);
			numberFormat.setParseBigDecimal(true);
			final BigDecimal parsed = (BigDecimal) numberFormat.parse(number, position);
			final String couldNoBeParsed = number.substring(position.getIndex());
			if ((parsed == null) || couldNoBeParsed.length() > 1) // trailing $, €, % symbols are ignored
				throw new ParseException(number + " is not a valid number!");

			return getValue(PrimitiveType.Decimal, column, parsed);
		}
	}

	private Object parseToDateTime(String column, ExcelCellData cellData) throws ParseException {
		Object value = cellData.getFormattedData();
		if (value == null)
			value = cellData.getRawData();
		if (cellData.getRawData() instanceof Double)
			value = cellData.getRawData();
		else if (cellData.getRawData() instanceof String && nrPattern.matcher((String) cellData.getRawData()).matches())
			value = Double.valueOf((String) cellData.getRawData());
		else if (value instanceof String) {
			if (nrPattern.matcher((String) value).matches())
				value = Double.valueOf((String) value);
			else if (cellData.getDisplayMask() != null) {
				String displayMask = cellData.getDisplayMask();
				if (displayMaskMap.containsKey(displayMask))
					settings.addDisplayMask(column, displayMaskMap.get(displayMask));
			} else if (settings.hasDefaultInputMask(column) != null) {
				settings.addDisplayMask(column, settings.getDefaultInputMask(column));
			} else if (!settings.hasValueParser(column))
				LOG_NODE.warn("Unable to parse the Date(" + value + ") in field: " + cellData.getColumnIndex());
		}

		if (value instanceof Double) {
			if (DateUtil.isValidExcelDate((Double) value)) {
				final TimeZone tz = settings.getTimeZoneForMember(column);
				value = DateUtil.getJavaDate((Double) value, tz);
			} else
				throw new ParseException("The value was not stored in excel as a valid date.");
		}

		return getValue(PrimitiveType.DateTime, column, value);
	}

	private final Pattern nrPattern = Pattern.compile("^\\d{0,6}(\\.\\d{1,})?$");
}
