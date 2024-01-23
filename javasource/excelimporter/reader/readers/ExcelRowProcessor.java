package excelimporter.reader.readers;

import com.mendix.replication.MendixReplicationException;

import java.util.Objects;

public interface ExcelRowProcessor {
    void processValues(ExcelRowProcessor.ExcelCellData[] values, int rowNow, int sheetNow) throws MendixReplicationException;
    void finish() throws MendixReplicationException;
    long getRowCounter();

    class ExcelCellData {
        private final int columnIndex;
        private final Object rawData;
        private final String displayMask;
        private final Object formattedData;

        public ExcelCellData(int columnIndex, Object rawData, Object formattedData) {
            this(columnIndex, rawData, formattedData, null);
        }

        public ExcelCellData(int columnIndex, Object rawData, Object formattedData, String displayMask) {
            this.columnIndex = columnIndex;
            this.rawData = rawData;
            this.formattedData = formattedData;
            this.displayMask = displayMask;
        }

        public String getDisplayMask() {
            return displayMask;
        }

        public int getColumnIndex() {
            return columnIndex;
        }

        public Object getFormattedData() {
            return formattedData;
        }

        public Object getRawData() {
            return rawData;
        }

        @Override
        public int hashCode() {
            return columnIndex + 31 * Objects.hash(rawData, displayMask, formattedData);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            ExcelCellData that = (ExcelCellData) o;
            return columnIndex == that.columnIndex &&
                    Objects.equals(rawData, that.rawData) &&
                    Objects.equals(displayMask, that.displayMask) &&
                    Objects.equals(formattedData, that.formattedData);
        }

        @Override
        public String toString() {
            return "ExcelCellData{" +
                    "colNr=" + columnIndex +
                    ", rawData=" + rawData +
                    ", displayMask='" + displayMask + '\'' +
                    ", formattedData=" + formattedData +
                    '}';
        }
    }
}
