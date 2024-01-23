package excelimporter.reader.readers.replication;

import com.mendix.replication.MendixReplicationException;
import com.mendix.replication.ReplicationSettings;

import java.util.TreeMap;

import com.mendix.systemwideinterfaces.core.IContext;

public class ExcelReplicationSettings extends ReplicationSettings {
	final private TreeMap<String, String> defaultInputMasks = new TreeMap<>();

	public ExcelReplicationSettings(IContext context, String objectType) throws MendixReplicationException {
		super(context, objectType, "ExcelImporter");
	}

	public Object hasDefaultInputMask(String column) {
		return defaultInputMasks.containsKey(column);
	}

	public String getDefaultInputMask(String column) {
		return defaultInputMasks.get(column);
	}

	public void addDefaultInputMask(String column, String mask) {
		if (mask != null)
			defaultInputMasks.put(column, mask);
	}
}
