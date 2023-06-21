package api.xl;

public final class XLCell<T> {
    private final String columnName;
    private final Integer rowNumber;
    private T value;

    /***
     * Creates an XLCell containing the row number, column name and a value of the proper Type.
     * @param columnName String
     * @param rowNumber Integer
     * @param value Object
     */
    public XLCell(String columnName, Integer rowNumber, T value) {
        this.columnName = columnName;
        this.rowNumber = rowNumber;
        this.value = value;

    }

    /***
     * Returns the column name
     * @return String
     */
    public String getColumnName() {
        return columnName;
    }
    /***
     * Returns the row number
     * @return Integer
     */
    public Integer getRowNumber() {
        return rowNumber;
    }

    /***
     * Return the cell value
     * @return Type
     */
    public T getValue() {
        return value;
    }

    public void setValue(T value) {
        this.value = value;
    }

    @Override
    public String toString() {
        if(value != null){
            return "XLCell{" +
                    "columnName='" + columnName + '\'' +
                    ", rowNumber=" + rowNumber +
                    ", value=" + value.toString() +
                    '}';
        }
        return "XLCell{" +
            "columnName='" + columnName + '\'' +
                    ", rowNumber=" + rowNumber +
                    ", value=" + "N/A" +
                    '}';
    }
}
