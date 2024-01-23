package excelimporter.reader.readers;

import java.util.Objects;

public class ExcelColumn {
	private final String caption;
	private final int columnIndex;

	public ExcelColumn(int columnIndex, String caption) {
		this.caption = caption;
		this.columnIndex = columnIndex;
	}

	public String getCaption() {
		return caption;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	@Override
	public int hashCode() {
		return columnIndex + 31 * Objects.hash(caption);
	}

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        final ExcelColumn that = (ExcelColumn) o;
        return columnIndex == that.columnIndex && Objects.equals(caption, that.caption);
    }

    @Override
    public String toString() {
        return "ExcelColumn{" + "caption='" + caption + '\'' + ", colNr=" + columnIndex + '}';
    }
}
