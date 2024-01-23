package excelimporter.reader.readers;

import com.mendix.core.Core;
import com.mendix.core.CoreException;
import com.mendix.logging.ILogNode;
import com.mendix.replication.AssociationConfig;
import com.mendix.replication.AssociationConfig.AssociationDataHandling;
import com.mendix.replication.ICustomValueParser;
import com.mendix.replication.MFValueParser;
import com.mendix.replication.ObjectConfig;
import com.mendix.replication.ObjectStatistics.Level;
import com.mendix.replication.ReplicationSettings.ChangeTracking;
import com.mendix.replication.ReplicationSettings.KeyType;
import com.mendix.replication.ReplicationSettings.ObjectSearchAction;
import com.mendix.replication.helpers.TimeMeasurement;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixIdentifier;
import com.mendix.systemwideinterfaces.core.IMendixObject;

import excelimporter.proxies.AdditionalProperties;
import excelimporter.proxies.Column;
import excelimporter.proxies.DataSource;
import excelimporter.proxies.ImportActions;
import excelimporter.proxies.MappingType;
import excelimporter.proxies.ReferenceDataHandling;
import excelimporter.proxies.ReferenceHandling;
import excelimporter.proxies.ReferenceHandlingEnum;
import excelimporter.proxies.ReferenceKeyType;
import excelimporter.proxies.RemoveIndicator;
import excelimporter.proxies.Template;
import excelimporter.reader.readers.replication.ExcelReplicationSettings;
import mxmodelreflection.proxies.MxObjectMember;
import mxmodelreflection.proxies.MxObjectReference;
import mxmodelreflection.proxies.MxObjectType;
import mxmodelreflection.proxies.PrimitiveTypes;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.util.RecordFormatException;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Read an Excel file can retrieve header information from by
 * 
 * @author J. van der Hoek - Mendix
 * @version $Id: ExcelXLSReader.java 9272 2009-05-11 09:19:47Z Jasper van der Hoek $
 */
public class ExcelReader {

	public static ILogNode logNode = Core.getLogger("ExcelXLSReader");

	private final TimeMeasurement timeMeasurement;
	private final IMendixObject templateObject;
	private final Map<String, Set<DocProperties>> docProperties = new HashMap<>();

	private ExcelReplicationSettings settings;
	private String descr;

	private static final String VALUE_STR = "=$value]";

	private static final String VALUE = "value";
	private long rowCount = 0L;

	public ExcelReader(IContext context, IMendixObject template) throws CoreException {
		if (template == null)
			throw new CoreException("No template was provided. Therefore the import could not be started.");

		timeMeasurement = new TimeMeasurement("ExcelReader-" + descr);
		templateObject = template;

		descr = templateObject.getValue(context, Template.MemberNames.Title.toString());
		if (descr != null)
			descr = templateObject.getValue(context, Template.MemberNames.Nr.toString()) + " - " + descr;
	}

	private static ExcelExtension getExcelExtension(IContext context, IMendixObject document) {
		final String name = ((String) document.getValue(context, "Name")).toLowerCase(Locale.ROOT);
		if (name.endsWith(".xls")) {
			return ExcelExtension.XLS;
		} else if (name.endsWith(".xlsx")) {
			return ExcelExtension.XLSX;
		} else if (name.endsWith(".xlsm")) {
			return ExcelExtension.XLSM;
		} else {
			return ExcelExtension.UNKNOWN;
		}
	}

	public List<ExcelColumn> getHeaders(IContext context, IMendixObject templateDocument) throws CoreException {
		if ((templateObject != null) && (context != null)) {
			StringBuilder sb = new StringBuilder("Starting XLS Headers import Template: ").append(descr);
			final String name = templateDocument.getValue(context, "Name");
			sb.append(" FileName: ").append(name != null ? name : "[unknown]");
			final ExcelExtension extension = getExcelExtension(context, templateDocument);
			logNode.info(sb.toString());
			final long importStartTime = System.nanoTime();

			File excelFile = null;

			try {
				// we maintain a 0-based rowNumber here.
				final int sheetIndex = (Integer) templateObject.getValue(context, Template.MemberNames.SheetIndex.toString()) - 1;
				if (sheetIndex < 0)
					throw new CoreException("The sheet-number must be >= 1");

				final int rowIndex = (Integer) templateObject.getValue(context, Template.MemberNames.HeaderRowNumber.toString()) - 1;
				if (rowIndex < 0)
					throw new CoreException("The header row-number must be >= 1");

				try (InputStream content = Core.getFileDocumentContent(context, templateDocument)) {
                    if (content == null)
                        throw new CoreException("No content found in templateDocument");
                }

				switch (extension) {
					case XLS:
					case XLSX:
					case XLSM:
						excelFile = getExcelFile(context, templateDocument);
						return ExcelHeaderReader.getHeaderColumns(excelFile, sheetIndex, rowIndex);
					case UNKNOWN:
						throw new CoreException("File extension is not an Excel extension ('.xls', '.xlsx' or '.xlsm')");
				}
			} catch (Exception e) {
				throw new CoreException("Document could not be read, because: " + e.getMessage(), e);
			} finally {
				if (excelFile != null) {
					try {
						excelFile.delete();
					} catch (Exception ignored) {
						logNode.info("Could not delete temp file.");
					} 
				}

				sb = new StringBuilder("Ready importing Headers ");
				sb.append((System.nanoTime() - importStartTime)/1000000).append(" ms");
				logNode.info(sb.toString());
			}
		}

		throw new CoreException("Template or context not set!"); // If this statement is reached no template was set
	}

	public long importData(IContext context, IMendixObject fileDocument, IMendixObject templateObject, IMendixObject parentObject)
			throws CoreException, ExcelImporterException {
		timeMeasurement.startPerformanceTest("Importing data");
		timeMeasurement.startPerformanceTest("Preparing all import settings");

		final IMendixIdentifier mxObjectType = templateObject.getValue(context, Template.MemberNames.Template_MxObjectType.toString());
		if (mxObjectType == null)
			throw new CoreException("There is no object type selected for the template");

		final String objectTypeName = Core.retrieveId(context, mxObjectType).getValue(context, MxObjectType.MemberNames.CompleteName.toString());

		settings = new ExcelReplicationSettings(context, objectTypeName);

		final ObjectConfig mainConfig = settings.getMainObjectConfig();
		switch (ImportActions.valueOf(templateObject.getValue(settings.getContext(), Template.MemberNames.ImportAction.toString()))) {
			case CreateObjects:
				mainConfig.setObjectSearchAction(ObjectSearchAction.CreateEverything);
				break;
			case SynchronizeObjects:
				mainConfig.setObjectSearchAction(ObjectSearchAction.FindCreate);
				break;
			case SynchronizeOnlyExisting:
				mainConfig.setObjectSearchAction(ObjectSearchAction.FindIgnore);
				break;
			case OnlyCreateNewObjects:
				mainConfig.setObjectSearchAction(ObjectSearchAction.OnlyCreateNewObjects);
				break;
		}
		setAdditionalProperties(settings.getContext(), settings, templateObject);

		final List<IMendixObject> columns = Core.createXPathQuery("//" + Column.getType() + "[" + Column.MemberNames.Column_Template + VALUE_STR)
				.setVariable(VALUE, templateObject.getId().toLong())
				.setAmount(Integer.MAX_VALUE)
				.setOffset(0)
				.setDepth(0)
				.addSort(Column.MemberNames.ColNumber.toString(), true)
				.execute(settings.getContext());
		for (IMendixObject columnObject : columns) {
			final DataSource dataSource = DataSource.valueOf(columnObject.getValue(settings.getContext(), Column.MemberNames.DataSource.toString()));
			final MappingType type = MappingType.valueOf(columnObject.getValue(settings.getContext(), Column.MemberNames.MappingType.toString()));

			final String fieldIdentifier;
			final DocProperties docProps;
			if (dataSource == DataSource.CellValue) {
				final Integer colNr = columnObject.getValue(settings.getContext(), Column.MemberNames.ColNumber.toString());
				fieldIdentifier = colNr.toString();
				docProps = null;
			}
			else {
				fieldIdentifier = String.valueOf(columnObject.getId().toLong());
				docProps = new DocProperties(dataSource, type, fieldIdentifier);
				if (dataSource == DataSource.StaticValue)
					docProps.setStaticStringValue(columnObject.getValue(context, Column.MemberNames.Text.toString()));
			}

			if (type == MappingType.Attribute) {
				final KeyType keyType = "Yes".equals(columnObject.getValue(settings.getContext(), Column.MemberNames.IsKey.toString()))
						? KeyType.ObjectKey : KeyType.NoKey;
				final boolean isCaseSensitive = "Yes".equals(columnObject.getValue(settings.getContext(), Column.MemberNames.CaseSensitive.toString()));

				final IMendixIdentifier microflow = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_Microflows.toString());
				final ICustomValueParser parser = (microflow != null)
					? new MFValueParser(settings.getContext(), Core.retrieveId(settings.getContext(), microflow))
					: null;

				final IMendixIdentifier memberId = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_MxObjectMember.toString());
				final IMendixObject memberObject = Core.retrieveId(settings.getContext(), memberId);
				final String attributeName = memberObject.getValue(settings.getContext(), MxObjectMember.MemberNames.AttributeName.toString());
				settings.addAttributeMapping(fieldIdentifier, attributeName, keyType, isCaseSensitive, parser);

				if (docProps != null) {
					if (!docProperties.containsKey(objectTypeName))
						docProperties.put(objectTypeName, new HashSet<>());

					docProperties.get(objectTypeName).add(docProps);
				}
			}

			else if (type == MappingType.Reference) {
				KeyType isKey = KeyType.NoKey;
				if (columnObject.getValue(settings.getContext(), Column.MemberNames.IsReferenceKey.toString()) != null
						&& !"".equals(columnObject.getValue(settings.getContext(), Column.MemberNames.IsReferenceKey.toString()))) {
					switch (ReferenceKeyType.valueOf(columnObject.getValue(settings.getContext(), Column.MemberNames.IsReferenceKey.toString()))) {
						case NoKey:
							isKey = KeyType.NoKey;
							break;
						case YesMainAndAssociatedObject:
							isKey = KeyType.AssociationAndObjectKey;
							break;
						case YesOnlyAssociatedObject:
							isKey = KeyType.AssociationKey;
							break;
						case YesOnlyMainObject:
							isKey = KeyType.ObjectKey;
							break;
					}
				}

				final boolean isCaseSensitive = "Yes".equals(columnObject.getValue(settings.getContext(), Column.MemberNames.CaseSensitive.toString()));

				final IMendixIdentifier microflow = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_Microflows.toString());
				final ICustomValueParser parser = (microflow != null)
						? new MFValueParser(settings.getContext(), Core.retrieveId(settings.getContext(), microflow))
						: null;

				final IMendixIdentifier memberId = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_MxObjectMember_Reference.toString());
				final IMendixObject memberObject = Core.retrieveId(settings.getContext(), memberId);
				final IMendixIdentifier associationId = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_MxObjectReference.toString());
				final IMendixObject associationObject = Core.retrieveId(settings.getContext(), associationId);
				final IMendixIdentifier objectTypeId = columnObject.getValue(settings.getContext(), Column.MemberNames.Column_MxObjectType_Reference.toString());
				final IMendixObject objectType = Core.retrieveId(settings.getContext(), objectTypeId);

				final String associationName = associationObject.getValue(settings.getContext(), MxObjectReference.MemberNames.CompleteName.toString());
				final String completeName = objectType.getValue(settings.getContext(), MxObjectType.MemberNames.CompleteName.toString());
				final String attributeName = memberObject.getValue(settings.getContext(), MxObjectMember.MemberNames.AttributeName.toString());
				settings.addAssociationMapping(fieldIdentifier, associationName, completeName, attributeName, isKey, isCaseSensitive, parser);

				if (docProps != null) {
					if (!docProperties.containsKey(associationName))
						docProperties.put(associationName, new HashSet<>());

					docProperties.get(associationName).add(docProps);
				}
			}

			if (PrimitiveTypes.DateTime.toString().equals(columnObject.getValue(context, Column.MemberNames.AttributeTypeEnum.toString()))) {
				settings.addDefaultInputMask(fieldIdentifier, columnObject.getValue(context, Column.MemberNames.InputMask.toString()));
			}
		}

		final List<IMendixObject> refHandlingList = Core.createXPathQuery("//" + ReferenceHandling.getType() + "[" + ReferenceHandling.MemberNames.ReferenceHandling_Template + VALUE_STR)
				.setVariable(VALUE, templateObject.getId().toLong())
				.execute(settings.getContext());
		for (IMendixObject object : refHandlingList) {
			final IMendixIdentifier refId = object.getValue(settings.getContext(), ReferenceHandling.MemberNames.ReferenceHandling_MxObjectReference.toString());
			final IMendixObject refObj = Core.retrieveId(settings.getContext(), refId);
			final String associationName = refObj.getValue(settings.getContext(), MxObjectReference.MemberNames.CompleteName.toString());
			final String selectReferenceHandling = object.getValue(settings.getContext(), ReferenceHandling.MemberNames.Handling.toString());

			final AssociationConfig config = settings.getAssociationConfig(associationName);

			switch (ReferenceHandlingEnum.valueOf(selectReferenceHandling)) {
				case FindCreate:
					config.setObjectSearchAction(ObjectSearchAction.FindCreate);
					break;
				case FindIgnore:
					config.setObjectSearchAction(ObjectSearchAction.FindIgnore);
					break;
				case CreateEverything:
					config.setObjectSearchAction(ObjectSearchAction.CreateEverything);
					break;
				case OnlyCreateNewObjects:
					config.setObjectSearchAction(ObjectSearchAction.OnlyCreateNewObjects);
					break;
			}

			final String refDataHandling = object.getValue(context, ReferenceHandling.MemberNames.DataHandling.toString());
			switch (ReferenceDataHandling.valueOf(refDataHandling)) {
				case Append:
					config.setAssociationDataHandling(AssociationDataHandling.Append);
					break;
				case Overwrite:
					config.setAssociationDataHandling(AssociationDataHandling.Overwrite);
					break;
			}

			config.setPrintNotFoundMessages(object.getValue(settings.getContext(), ReferenceHandling.MemberNames.PrintNotFoundMessages.toString()));
			config.setCommitUnchangedObjects(object.getValue(settings.getContext(), ReferenceHandling.MemberNames.CommitUnchangedObjects.toString()));
			config.setIgnoreEmptyKeys(object.getValue(context, ReferenceHandling.MemberNames.IgnoreEmptyKeys.toString()));
		}

		final IMendixIdentifier parentAssociationId = templateObject.getValue(settings.getContext(), Template.MemberNames.Template_MxObjectReference_ParentAssociation.toString());
		if (parentAssociationId != null) {
			final IMendixObject parentAssociation = Core.retrieveId(settings.getContext(), parentAssociationId);
			final String associationName = parentAssociation.getValue(settings.getContext(), MxObjectReference.MemberNames.CompleteName.toString());

			settings.setParentObject(parentObject, associationName);
		}
		timeMeasurement.endPerformanceTest("Preparing all import settings");
		logNode.info("Starting XLS import Template: " + descr);

		final long importStartTime = System.nanoTime();

		File excelFile = null;
		try {
			// We store sheetIndex and startRowIndex as a zero-based number; start importing at this sheet / row
			final int sheetIndex = (Integer) templateObject.getValue(settings.getContext(), Template.MemberNames.SheetIndex.toString()) - 1;
			if (sheetIndex < 0)
				throw new CoreException("The sheet-number must be >= 1");

			final int startRowIndex = (Integer) templateObject.getValue(settings.getContext(), Template.MemberNames.FirstDataRowNumber.toString()) - 1;
			if (startRowIndex < 0)
				throw new CoreException("The row-number must be >= 1");

			switch (getExcelExtension(settings.getContext(), fileDocument)) {
				case XLS:
				case XLSX:
				case XLSM:
					final var xpath = String.format("count(//%s[%s = %d])", Column.entityName, Column.MemberNames.Column_Template, templateObject.getId().toLong());					logNode.debug(xpath);

					final long columnCount = Core.createXPathQuery(xpath)
							.executeAggregateLong(settings.getContext());
					logNode.debug("nrOfCols: " + columnCount);

					excelFile = getExcelFile(settings.getContext(), fileDocument);
					rowCount = ExcelDataReader.readData(excelFile.getAbsolutePath(), sheetIndex, startRowIndex, new ExcelRowProcessorImpl(getSettings(), getDocPropertiesMapping()), getSettings()::aliasIsMapped);
					break;
				case UNKNOWN:
					throw new CoreException("File extension is not an Excel extension ('.xls', '.xlsx' or 'xlsm').");
			}

			logNode.info("Successfully finished importing " + descr + " in " + ((System.nanoTime() - importStartTime) / 1000000) + " ms");
		} catch (OLE2NotOfficeXmlFileException e) {
			logNode.info("Error while importing " + descr + " " + ((System.nanoTime() - importStartTime) / 1000000) + " ms, because: " + e.getMessage());
			throw new ExcelImporterException("Document could not be imported because this file is an XLS and not an XLSX file. Please make sure the file is valid and has the correct extension.");
		} catch (NotOfficeXmlFileException | NotOLE2FileException e) {
			logNode.info("Error while importing " + descr + " " + ((System.nanoTime() - importStartTime) / 1000000) + " ms, because: " + e.getMessage());
			throw new ExcelImporterException("Document could not be imported because this file is not XLS or XLSX. Please make sure the file is valid and has the correct extension.");
		} catch (RecordFormatException e) {
			logNode.info("Error while importing " + descr + " " + ((System.nanoTime() - importStartTime) / 1000000) + " ms, because: " + e.getMessage());
			throw new ExcelImporterException("Document could not be imported because one of its cell values is invalid or cannot be read.");
		} catch (EncryptedDocumentException e) {
			logNode.info("Error while importing " + descr + " " + ((System.nanoTime() - importStartTime) / 1000000) + " ms, because: " + e.getMessage());
			throw new ExcelImporterException("Document could not be imported because it is encrypted.");
		} catch (Exception e) {
			logNode.info("Error while importing " + descr + " " + ((System.nanoTime() - importStartTime) / 1000000) + " ms, because: " + e.getMessage());
			throw new CoreException("Document could not be imported, because: " + e.getMessage(), e);
		} finally {
			docProperties.clear();
			settings.clear();
			if (excelFile != null) {
				try {
					excelFile.delete();
				} catch (final Exception ignored) {
					logNode.info("Could not delete temp file.");
				}
			}

			timeMeasurement.endPerformanceTest("Importing data");
		}
		return rowCount;
	}

	private static File getExcelFile(IContext context, IMendixObject file) throws IOException {
		final File f = new File(Core.getConfiguration().getTempPath().getAbsolutePath() + "/Mendix_ExcelImporter_" + file.getId().toLong(), "");
		try (InputStream inputstream = Core.getFileDocumentContent(context, file); OutputStream outputstream = new FileOutputStream(f)) {
			final byte[] buffer = new byte[4 * 1024];
			int length;
			while ((length = inputstream.read(buffer)) > 0) {
				outputstream.write(buffer, 0, length);
			}
		}
		return f;
	}

	/**
	 * Set the additional properties in the settings object
	 * This is done by retrieving the associated AdditionalProperties object from the Template.
	 * 
	 * @param context
	 * @param settings
	 * @param templateObject
	 * @throws CoreException
	 */
	private static void setAdditionalProperties(IContext context, ExcelReplicationSettings settings, IMendixObject templateObject) throws CoreException {
		final IMendixIdentifier addPropertyId = templateObject.getValue(context, Template.MemberNames.Template_AdditionalProperties.toString());
		if (addPropertyId != null) {
			final IMendixObject addProperties = Core.retrieveId(context, addPropertyId);
			if (addProperties != null) {
				if (addProperties.getValue(context, AdditionalProperties.MemberNames.PrintStatisticsMessages.toString()) == null)
					settings.setStatisticsLevel(Level.AllStatistics);
				else {
					switch (excelimporter.proxies.StatisticsLevel.valueOf(addProperties.getValue(context, AdditionalProperties.MemberNames.PrintStatisticsMessages.toString()))) {
						case OnlyFinalStatistics:
							settings.setStatisticsLevel(Level.OnlyFinalStatistics);
							break;
						case NoStatistics:
							settings.setStatisticsLevel(Level.NoStatistics);
							break;
						case AllStatistics:
						default:
							settings.setStatisticsLevel(Level.AllStatistics);
							break;
					}
				}
				settings.getMainObjectConfig().setPrintNotFoundMessages(addProperties.getValue(context, AdditionalProperties.MemberNames.PrintNotFoundMessages_MainObject.toString()));
				settings.setIgnoreEmptyKeys(addProperties.getValue(context, AdditionalProperties.MemberNames.IgnoreEmptyKeys.toString()));

				settings.getMainObjectConfig().setCommitUnchangedObjects(addProperties.getValue(context, AdditionalProperties.MemberNames.CommitUnchangedObjects_MainObject.toString()));

				final String indicator = addProperties.getValue(context, AdditionalProperties.MemberNames.RemoveUnsyncedObjects.toString());
				final ObjectConfig mainConfig = settings.getMainObjectConfig();
				switch (RemoveIndicator.valueOf(indicator)) {
					case RemoveUnchangedObjects:
						final IMendixIdentifier removeId = addProperties.getValue(context, AdditionalProperties.MemberNames.AdditionalProperties_MxObjectMember_RemoveIndicator.toString());
						if (removeId != null) {
							final IMendixObject removeObject = Core.retrieveId(context, removeId);
							final String attributeName = removeObject.getValue(context, MxObjectMember.MemberNames.AttributeName.toString());
							mainConfig.removeUnusedObjects(ChangeTracking.RemoveUnchangedObjects, attributeName);
						}
						break;
					case TrackChanges:
						final IMendixIdentifier trackId = addProperties.getValue(context, AdditionalProperties.MemberNames.AdditionalProperties_MxObjectMember_RemoveIndicator.toString());
						if (trackId != null) {
							final IMendixObject trackObject = Core.retrieveId(context, trackId);
							final String attributeName = trackObject.getValue(context, MxObjectMember.MemberNames.AttributeName.toString());
							mainConfig.removeUnusedObjects(ChangeTracking.TrackChanges, attributeName);
						}
						break;
					case Nothing:
					default:
						mainConfig.removeUnusedObjects(ChangeTracking.Nothing, null);
						break;
				}

				if (addProperties.hasMember("RetrieveOQL_Limit")) {
					final Integer limit = addProperties.getValue(context, "RetrieveOQL_Limit");
					if (limit != null) {
						logNode.debug("Changing retrieve limit to: " + limit);
						settings.getConfiguration().RetrieveOQL_Limit = limit;
					}
				}
			}
		}
	}

	public ExcelReplicationSettings getSettings() {
		return settings;
	}

	public Map<String, Set<DocProperties>> getDocPropertiesMapping() {
		return docProperties ;
	}
}
