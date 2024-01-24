package xlsreport.report.export;

import com.mendix.core.Core;
import com.mendix.systemwideinterfaces.core.IContext;
import com.mendix.systemwideinterfaces.core.IMendixObject;
import com.opencsv.CSVWriter;
import com.opencsv.ICSVWriter;
import org.joda.time.DateTime;
import org.joda.time.DateTimeZone;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;
import system.proxies.FileDocument;
import xlsreport.proxies.MxSheet;
import xlsreport.proxies.MxTemplate;
import xlsreport.report.data.ColumnPreset;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

public class ExportCSV  extends Export
{
	private ByteArrayOutputStream baos;
	private CSVWriter writer;
	DateTimeFormatter fmt;

	public ExportCSV(IContext context, MxTemplate template, IMendixObject inputObject)
	{
		super(context,inputObject);
		char separator = ',';
		if(template.getCSVSeparator() != null)
		{
			switch(template.getCSVSeparator())
			{
				case Comma:
					separator = ',';
					break;
				case Semicolon:
					separator = ';';
					break;
				case Tab:
					separator = '\t';
					break;
			}
		}
		String format = Export.getDatePresentation(template);
		fmt = DateTimeFormat.forPattern(format);
		char quotecharacter;
		if(template.getQuotationCharacter() == null || template.getQuotationCharacter().equals("")){
			quotecharacter = ICSVWriter.NO_QUOTE_CHARACTER;
		}
		else {
			quotecharacter = template.getQuotationCharacter().charAt(0);
		}
		// Create the OpenCSV writer
		this.baos = new ByteArrayOutputStream();

		this.writer = new CSVWriter(new OutputStreamWriter(baos), separator, quotecharacter, '"', "\r\n");
	}

	@Override
	public void buildExportFile(MxSheet mxSheet, List<ColumnPreset> mxColumnList, Object[][] table) throws Exception
	{
		String[] values = new String[mxColumnList.size()];
		for(int i = 0; i < mxColumnList.size(); i++)
		{
			values[i] = mxColumnList.get(i).getName();
		}
		writer.writeNext(values);
		
		//SimpleDateFormat format = new SimpleDateFormat(this.dateFormat);		
		for(int i = 0; i < table.length; i++)
		{
			for(int e = 0; e < table[i].length; e++)
			{
				Object value = table[i][e];
				if(value instanceof Date)
				{
					DateTime dateTime = new DateTime(value);
					if(mxColumnList.get(e).isShouldLocalizeDate())
					{
						value = dateTime.withZone(DateTimeZone.forTimeZone(context.getSession().getTimeZone())).toLocalDateTime().toString(fmt);
					}
					else
					{
						value = dateTime.withZone(DateTimeZone.UTC).toLocalDateTime().toString(fmt);
					}
				}
				if (value == null) { 
					values[e] = null; 
				} 
				else { 
					values[e] = String.valueOf(value); 
				}		
			}			
			writer.writeNext(values);			
		}
		writer.flush();
	}	
	
	@Override
	public void writeData(FileDocument outputDocument) throws Exception
	{
		try (InputStream inputStream = new ByteArrayInputStream(baos.toByteArray()))
		{
			Core.storeFileDocumentContent(context, outputDocument.getMendixObject(), inputStream);
		}
	}

	public void close() throws Exception
	{
		if (this.baos != null)
			this.baos.close();
		if (this.writer != null)
			this.writer.close();
	}
}
