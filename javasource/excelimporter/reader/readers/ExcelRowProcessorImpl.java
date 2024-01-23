package excelimporter.reader.readers;

import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import com.mendix.replication.MetaInfo;
import com.mendix.replication.MetaInfoObject;
import com.mendix.replication.MendixReplicationException;
import com.mendix.systemwideinterfaces.core.meta.IMetaPrimitive.PrimitiveType;

import excelimporter.reader.readers.replication.ExcelReplicationSettings;
import excelimporter.reader.readers.replication.ExcelValueParser;

public class ExcelRowProcessorImpl implements ExcelRowProcessor {
	private final ExcelValueParser valueParser;
	private final MetaInfo info;
	private final ExcelReplicationSettings settings;
	private final Map<String, Set<DocProperties>> docProps;
	private final boolean hasDocProps;
	private long rowCounter;

	public ExcelRowProcessorImpl(ExcelReplicationSettings settings, Map<String, Set<DocProperties>> docProps) throws MendixReplicationException {
	    this.settings = settings;
		this.valueParser = new ExcelValueParser(settings);
		this.info = new MetaInfo(settings, valueParser, "XLSReader");
		this.docProps = docProps;
		this.hasDocProps = docProps.size() > 0;

		this.rowCounter = 0;
	}

	public void processValues(ExcelRowProcessor.ExcelCellData[] values, int rowNow, int sheetNow) throws MendixReplicationException {
		final String objectKey = valueParser.buildObjectKey(values, settings.getMainObjectConfig());
		if (ExcelReader.logNode.isTraceEnabled())
			ExcelReader.logNode.trace("Start processing excel row: " + rowCounter + " found: " + values.length + " columns to process. Using ObjectKey: " + objectKey);

		final Map<String, Long> prevObject = (hasDocProps) ? new HashMap<>() : null;
		rowNow++;
		sheetNow++;

		final List<String> columnAliases = settings.getAliasList();
		final List<String> docPropColumnAliases = !hasDocProps ? Collections.emptyList()
				: docProps.values().stream().flatMap(set -> set.stream().map(DocProperties::getColumnAlias)).collect(Collectors.toList());

		for (String alias : columnAliases) {
			if (!docPropColumnAliases.contains(alias)) {
				final PrimitiveType type = settings.getMemberType(alias);
				final Object processedValue = valueParser.getValueFromDataSet(alias, type, values);

				final String id;
				final MetaInfoObject miObject;
				if (settings.treatFieldAsReference(alias)) {
					miObject = info.setAssociationValue(objectKey, alias, processedValue);
					id = settings.getAssociationNameByAlias(alias);
				}
				else if (settings.treatFieldAsReferenceSet(alias)) {
					miObject = info.addAssociationValue(objectKey, alias, processedValue);
					id = settings.getAssociationNameByAlias(alias);
				}
				else {
					miObject = info.addValue(objectKey, alias, processedValue);
					id = settings.getMainObjectConfig().getObjectType();
				}
				final Long columnObjectID = (miObject == null) ? null : miObject.getId();

				if (hasDocProps && docProps.containsKey(id)) {
					if (!prevObject.containsKey(id) || columnObjectID != prevObject.get(id)) {
						prevObject.put(id, columnObjectID);

						for (DocProperties props : docProps.get(id)) {
							Object value = null;
							switch (props.getDataSource()) {
							case DocumentPropertyRowNr:
								value = rowNow;
								break;
							case DocumentPropertySheetNr:
								value = sheetNow;
								break;
							case StaticValue:
								value = props.getStaticStringValue();
								break;
							case CellValue:
								break;
							}

							switch (props.getMappingType()) {
							case Attribute:
								info.addValue(objectKey, props.getColumnAlias(), value);
								break;
							case Reference:
								info.addAssociationValue(objectKey, props.getColumnAlias(), value);
								break;
							default:
								break;
							}
						}
					}
				}
			}
		}

		rowCounter++;
		resetValuesArray(values);
	}

	private static void resetValuesArray(Object[] values) {
        Arrays.fill(values, null);
	}

	public void finish() throws MendixReplicationException {
		info.finish();
		info.clear();
	}

	public long getRowCounter() {
		return rowCounter;
	}
}
