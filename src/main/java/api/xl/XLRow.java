package api.xl;

import java.util.Comparator;
import java.util.LinkedList;
import java.util.List;

public final class XLRow {
    private final int rowNum;
    private final List<XLCell<?>> row = new LinkedList<>();

    public XLRow(int rowNum) {
        this.rowNum = rowNum;
    }

    public void addXlCell(XLCell<?> cell){
        row.add(cell);
    }

    /**
     * Size of the list of cells in the row.
     * @return An int value of the size of the cell's list.
     */
    public int getCellCount(){
        return row.size();
    }
    public XLCell<?> getLastCell(){
        return row.get(row.size()-1);
    }
    public XLCell<?> getCell(int idx){
        return row.get(idx);
    }
    public LinkedList<XLCell<?>> asList() {
        return new LinkedList<>(row);
    }
    public void sort(){
        row.sort(Comparator.comparing(XLCell::getColumnName));
    }
    public Integer getRowNum() {
        return rowNum;
    }
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        for (XLCell<?> c: row) {
            sb.append(c.toString()).append("\n");
        }
        return sb.toString();
    }
}
