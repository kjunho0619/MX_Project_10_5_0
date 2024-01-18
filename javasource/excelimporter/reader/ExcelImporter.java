package excelimporter.reader;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map.Entry;

import system.proxies.FileDocument;

import com.mendix.core.Core;
import com.mendix.core.CoreException;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixObject;

import excelimporter.proxies.Column;
import excelimporter.proxies.Template;
import excelimporter.reader.readers.ExcelColumn;
import excelimporter.reader.readers.ExcelReader;

/**
 * @author J. van der Hoek - Mendix
 * @version $Id: ExcelImporter.java 9272 2009-05-11 09:19:47Z Jasper van der Hoek $
 */
public class ExcelImporter {

	/**
	 * Get all columns from the TemplateDocument and store them in the Template
	 * @param context
	 * @param template
	 * @throws CoreException
	 */
	public static void getColumns(IContext context, IMendixObject template, IMendixObject excelDocument) throws CoreException {
		if (excelDocument == null)
			throw new CoreException("You must select a file document before the columns can be extracted.");
		if (!excelDocument.isInstanceOf(FileDocument.getType()))
			throw new CoreException("The exceldocument object must be a sublclass of System.FileDocument.");

		final String path = excelDocument.getValue(context, "Name");
		if (path != null) {
			// We need to check a radiobutton in the template, not the file extension.

			// Determine file format for header extraction
			final int lastdot = path.lastIndexOf(".");
			if (lastdot < 0)
				throw new CoreException("Found file has no extension to derive format from.");

			final String extension = path.substring(lastdot).toLowerCase(Locale.ROOT);

			if (".xls".equals(extension) || ".xlsx".equals(extension) || ".xlsm".equals(extension)) {
				final ExcelReader excelXLSReader = new ExcelReader(context, template);
				final List<ExcelColumn> columnList = excelXLSReader.getHeaders(context, excelDocument);

				if (columnList.size() == 0)
					throw new CoreException("No column information could be found in document.");

				final List<IMendixObject> columns = Core.createXPathQuery("//" + Column.getType() + "[" + Column.MemberNames.Column_Template + "=$value]")
						.setVariable("value", template.getId().toLong())
						.execute(context);
				final HashMap<Integer, IMendixObject> existingColumns = new HashMap<>(columns.size());
				for (IMendixObject col : columns) {
					existingColumns.put(col.getValue(context, Column.MemberNames.ColNumber.toString()), col);
				}

				for (ExcelColumn excelColumn : columnList) {
					final Integer columnIndex = excelColumn.getColumnIndex();
					final IMendixObject column;
					if (existingColumns.containsKey(columnIndex)) {
						column = existingColumns.get(columnIndex);
						existingColumns.remove(columnIndex);
					} else {
						column = Core.instantiate(context, Column.getType());
						column.setValue(context, Column.MemberNames.ColNumber.toString(), columnIndex);
						column.setValue(context, Column.MemberNames.Column_Template.toString(), template.getId());
					}

					column.setValue(context, Column.MemberNames.Column_MxObjectType.toString(), template.getValue(context, Template.MemberNames.Template_MxObjectType.toString()));
					column.setValue(context, Column.MemberNames.Text.toString(), excelColumn.getCaption());
					Core.commit(context, column);
				}

				for (Entry<Integer, IMendixObject> entry : existingColumns.entrySet()) {
					Core.delete(context, entry.getValue());
				}
			}
			else
				throw new CoreException("The extension: " + extension + " is not supported.");
		}
		else
			throw new CoreException("No valid filepath found from template");
	}

	/**
	 * Import the file document using the template from the parameter
	 *
	 * @param context
	 * @param template
	 * @param templateDocument
	 * @param parentObject the object to which the objects to be imported, must be associated.
	 */
	public static long startImport(IContext context, IMendixObject template, IMendixObject templateDocument, IMendixObject parentObject) throws CoreException {
		if (template == null)
			throw new CoreException("Template not found!");

		try (InputStream content = Core.getFileDocumentContent(context, templateDocument);) {
			if (content != null) {
				final String fileName = templateDocument.getValue(context, "Name");
				if (fileName != null) {
					// Determine file format for header extraction
					final int lastdot = fileName.lastIndexOf(".");
					if (lastdot < 0)
						throw new CoreException("Found file has no extension to derive format from.");

					final String extension = fileName.substring(lastdot).toLowerCase(Locale.ROOT);

					if (".xls".equalsIgnoreCase(extension) || ".xlsx".equalsIgnoreCase(extension) || ".xlsm".equalsIgnoreCase(extension)) {
						final ExcelReader reader = new ExcelReader(context, template);
						return reader.importData(context, templateDocument, template, parentObject);
					} else {
						throw new CoreException("The extension: " + extension + " is not supported.");
					}

				} else
					throw new CoreException("No valid filepath found from template");
			}
		} catch (Exception e) {
			throw new CoreException(e);
		}
		return 0L;
	}
}
