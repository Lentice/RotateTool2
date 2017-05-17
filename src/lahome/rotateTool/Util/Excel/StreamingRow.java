package lahome.rotateTool.Util.Excel;

import org.apache.poi.ss.usermodel.*;

import java.util.Iterator;

@SuppressWarnings("WeakerAccess")
public class StreamingRow implements Row {
    private int rowIndex;
    private int maxColumnIndex = -1;
    //private ArrayList<Cell> cellMap = new ArrayList<Cell>(255);
    private StreamingCell[] cellArray = new StreamingCell[127];

    public StreamingRow(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public void clear() {
        int minSize = Math.min(maxColumnIndex, cellArray.length);
        for (int i = 0; i <= minSize; i++) {
            if (cellArray[i] != null) {
                cellArray[i].setRawContents(null);
            }
        }
        maxColumnIndex = -1;
    }

    public StreamingCell createCell(int columnIndex) {
        if (cellArray[columnIndex] == null) {
            cellArray[columnIndex] = new StreamingCell(this, columnIndex);
        }

        maxColumnIndex = Math.max(columnIndex, maxColumnIndex);
        return cellArray[columnIndex];
    }
//}
//    public void addCell(int columnIndex, Cell cell) {
//        if (cellArray[columnIndex] == null) {
//            cell = new StreamingCell(this, columnIndex);
//
//        cellArray[columnIndex] = cell;
//        maxColumnIndex = Math.max(columnIndex, maxColumnIndex);
//    }

    /**
     * Get row number this row represents
     *
     * @return the row number (0 based)
     */
    public int getRowNum() {
        return rowIndex;
    }


    @Override
    public void setRowNum(int rowNum) {
        this.rowIndex = rowNum;
    }

    /**
     * Get the cell representing a given column (logical cell) 0-based.  If you
     * ask for a cell that is not defined, you get a null.
     *
     * @param cellColumnIndex 0 based column number
     * @return Cell representing that column or null if undefined.
     */
    public Cell getCell(int cellColumnIndex) {
        return cellArray[cellColumnIndex];
    }

    public int getMaxColumnIndex() {
        return maxColumnIndex;
    }


    /* Not supported */
    @Override
    public int getPhysicalNumberOfCells() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getFirstCellNum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Cell> cellIterator() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Iterator<Cell> iterator() {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getLastCellNum() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean getZeroHeight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public Cell getCell(int cellnum, MissingCellPolicy policy) {
        throw new UnsupportedOperationException();
    }

    @Override
    @Deprecated
    public Cell createCell(int column, int type) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Cell createCell(int i, CellType cellType) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void removeCell(Cell cell) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHeight(short height) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setZeroHeight(boolean zHeight) {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setHeightInPoints(float height) {
        throw new UnsupportedOperationException();
    }

    @Override
    public short getHeight() {
        throw new UnsupportedOperationException();
    }

    @Override
    public float getHeightInPoints() {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isFormatted() {
        throw new UnsupportedOperationException();
    }

    @Override
    public CellStyle getRowStyle() {
        throw new UnsupportedOperationException();
    }

    @Override
    public void setRowStyle(CellStyle style) {
        throw new UnsupportedOperationException();
    }

    @Override
    public Sheet getSheet() {
        throw new UnsupportedOperationException();
    }

    @Override
    public int getOutlineLevel() {
        throw new UnsupportedOperationException();
    }

}
