package lahome.rotateTool.Util;

import com.monitorjbl.xlsx.StreamingReader;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.*;
import java.util.Date;

public class ExcelParser {
    private static final Logger log = LogManager.getLogger(ExcelParser.class.getName());

    private static final int sheetIdx = 0;

    private RotateCollection collection;

    private InputStream is;
    private Workbook wb;

    public ExcelParser(RotateCollection collection) {
        this.collection = collection;
    }


    private String getKitName(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        String kitName = cellGetString(cell).trim();
        if (kitName.contains("空白"))
            kitName = "";

        return kitName;
    }

    private String getPartNumber(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim();
    }

    private int getExpectQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private String getPo(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim();
    }

    private int getStockQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private String getLot(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim();
    }

    private String getDc(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }
        return cellGetString(cell).trim();
    }

    private int getMyQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            //throw new Exception(String.format("myQty is null at row %d", row.getRowNum()));
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private String getRatio(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim();
    }

    private int getApplySet(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private String getRemark(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim();
    }

    private String getGrDateQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }
        return cellGetDate(cell);
    }

    private int getGrQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            //throw new Exception(String.format("myQty is null at row %d", row.getRowNum()));
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private Sheet getSheet(File file) {

        if (!file.exists()) {
            log.warn(String.format("%s(%d) file is not exists",
                    Thread.currentThread().getStackTrace()[1].getMethodName(),
                    Thread.currentThread().getStackTrace()[1].getLineNumber()));
            return null;
        }

        try {
            is = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            log.error(String.format("File %s opened failed", file.getPath()));
            return null;
        }

        wb = StreamingReader.builder()
                .rowCacheSize(10)     // number of rows to keep in memory (defaults to 10)
                .bufferSize(1024)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)
        if (wb == null) {
            log.warn(String.format("%s(%d) wb is null",
                    Thread.currentThread().getStackTrace()[1].getMethodName(),
                    Thread.currentThread().getStackTrace()[1].getLineNumber()));
            return null;
        }

        return wb.getSheetAt(sheetIdx);
    }

    private void close() {

        try {
            if (is != null) {
                is.close();
            }
            if (wb != null) {
                wb.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private String cellGetString(Cell cell) {
        int cellType = cell.getCellType();
        if (cellType == Cell.CELL_TYPE_FORMULA) {
            switch (cell.getCachedFormulaResultType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    long value = (long) cell.getNumericCellValue();
                    return Long.toString(value);
                case Cell.CELL_TYPE_STRING:
                    return cell.getStringCellValue();
            }
        } else if (cellType == Cell.CELL_TYPE_STRING) {
            return cell.getStringCellValue();
        } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
            long value = (long) cell.getNumericCellValue();
            return Long.toString(value);
        } else if (cellType == Cell.CELL_TYPE_BLANK) {
            return "";
        }

        log.warn(String.format("%s(%d) cell(%d,%d) is unknown type",
                Thread.currentThread().getStackTrace()[1].getMethodName(),
                Thread.currentThread().getStackTrace()[1].getLineNumber(),
                cell.getRowIndex(), cell.getColumnIndex()));

        return "";
    }

    private double cellGetValue(Cell cell) {
        try {
            int cellType = cell.getCellType();
            if (cellType == Cell.CELL_TYPE_FORMULA) {
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return cell.getNumericCellValue();
                    case Cell.CELL_TYPE_STRING:
                        return Double.parseDouble(numTrim(cell.getStringCellValue().trim()));
                }
            } else if (cellType == Cell.CELL_TYPE_STRING) {
                String str = numTrim(cell.getStringCellValue().trim());
                if (str.isEmpty())
                    return 0;
                else
                    return Double.valueOf(str);
            } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cellType == Cell.CELL_TYPE_BLANK) {
                return 0;
            }

            log.warn(String.format("%s(%d) cell(%d,%d) is unknown type",
                    Thread.currentThread().getStackTrace()[1].getMethodName(),
                    Thread.currentThread().getStackTrace()[1].getLineNumber(),
                    cell.getRowIndex(), cell.getColumnIndex()));

            return 0;
        } catch (Exception e) {
            log.error("failed!", e);
            return 0;
        }
    }

    private String cellGetDate(Cell cell) {
        try {
            Date date = cell.getDateCellValue();
            return DateUtil.format(date);
        } catch (Exception e) {
            log.error("failed!", e);
            return null;
        }
    }

    private String numTrim(String s) {
        StringBuilder sb = new StringBuilder(s);

        int i = 0;
        while (i < sb.length()) {
            if (!isNumber(sb.charAt(i)))
                sb.deleteCharAt(i);
            else
                i++;
        }
        return sb.toString();
    }

    private boolean isNumber(char c) {
        return (c >= '0' && c <= '9');
    }


    public void loadRotateExcel(File file, int firstDataRow, String kitColStr, String partColStr,
                                String pmQtyColStr, String ratioColStr, String applySetColStr,
                                String remarkColStr) {

        firstDataRow = firstDataRow - 1;
        int kitColumn = CellReference.convertColStringToIndex(kitColStr.toUpperCase());
        int partColumn = CellReference.convertColStringToIndex(partColStr.toLowerCase());
        int pmQtyColumn = CellReference.convertColStringToIndex(pmQtyColStr.toUpperCase());
        int ratioColumn = CellReference.convertColStringToIndex(ratioColStr.toUpperCase());
        int applySetColumn = CellReference.convertColStringToIndex(applySetColStr.toUpperCase());
        int remarkColumn = CellReference.convertColStringToIndex(remarkColStr.toUpperCase());

        int maxColumnUsed = Math.max(Math.max(kitColumn, partColumn), pmQtyColumn);

        Sheet sheet = getSheet(file);
        if (sheet == null) {
            log.error("open sheet failed");
            return;
        }

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < firstDataRow)
                continue;

            int columnCount = row.getPhysicalNumberOfCells();
            if (columnCount < maxColumnUsed)
                continue;

            try {
                String kitName = getKitName(row, kitColumn);
                String partNum = getPartNumber(row, partColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                int pmQty = getExpectQty(row, pmQtyColumn);

                if (pmQty == 0)
                    continue;

                String ratio = getRatio(row, ratioColumn);
                int applySet = getApplySet(row, applySetColumn);
                String remark = getRemark(row, remarkColumn);

                collection.addRotate(new RotateItem(rowNum, kitName, partNum, pmQty, ratio, applySet, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close();
    }

    public void loadStockExcel(File file, int firstDataRow, String kitColStr, String partColStr,
                               String poColStr, String stockQtyColStr, String lotColStr, String dcColStr,
                               String myQtyColStr, String remarkColStr) {
        firstDataRow = firstDataRow - 1;

        int kitColumn = CellReference.convertColStringToIndex(kitColStr);
        int partColumn = CellReference.convertColStringToIndex(partColStr);
        int poColumn = CellReference.convertColStringToIndex(poColStr);
        int stockQtyColumn = CellReference.convertColStringToIndex(stockQtyColStr);
        int lotColumn = CellReference.convertColStringToIndex(lotColStr);
        int dcColumn = CellReference.convertColStringToIndex(dcColStr);
        int myQtyColumn = CellReference.convertColStringToIndex(myQtyColStr);
        int remarkColumn = CellReference.convertColStringToIndex(remarkColStr);

        int maxColumnUsed = Math.max(kitColumn, Math.max(partColumn, Math.max(poColumn, Math.max(stockQtyColumn,
                Math.max(stockQtyColumn, Math.max(dcColumn, myQtyColumn))))));

        Sheet sheet = getSheet(file);
        if (sheet == null) {
            log.error("open sheet failed");
            return;
        }

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < firstDataRow)
                continue;

            int columnCount = row.getPhysicalNumberOfCells();
            if (columnCount < maxColumnUsed)
                continue;

            try {
                String kitName = getKitName(row, kitColumn);
                String partNum = getPartNumber(row, partColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.isRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, poColumn);
                if (po.isEmpty())
                    continue;

                int stockQty = getStockQty(row, stockQtyColumn);
                String lot = getLot(row, lotColumn);
                String dc = getDc(row, dcColumn);
                int myQty = getMyQty(row, myQtyColumn);
                String remark = getRemark(row, remarkColumn);

                collection.addStock(kitName, partNum,
                        new StockItem(rowNum, po, stockQty, lot, dc, myQty, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close();
    }

    public void loadPurchaseExcel(File file, int firstDataRow, String kitColStr, String partColStr,
                                  String poColStr, String grDateColStr, String grQtyColStr,
                                  String myQtyColStr, String setColStr, String remarkColStr) {
        firstDataRow = firstDataRow - 1;
        int kitColumn = CellReference.convertColStringToIndex(kitColStr);
        int partColumn = CellReference.convertColStringToIndex(partColStr);
        int poColumn = CellReference.convertColStringToIndex(poColStr);
        int grDateColumn = CellReference.convertColStringToIndex(grDateColStr);
        int grQtyColumn = CellReference.convertColStringToIndex(grQtyColStr);
        int myQtyColumn = CellReference.convertColStringToIndex(myQtyColStr);
        int setColumn = CellReference.convertColStringToIndex(setColStr);
        int remarkColumn = CellReference.convertColStringToIndex(remarkColStr);


        int maxColumnUsed = Math.max(kitColumn, Math.max(partColumn, Math.max(poColumn, Math.max(grDateColumn,
                grQtyColumn))));

        Sheet sheet = getSheet(file);
        if (sheet == null) {
            log.error("open sheet failed");
            return;
        }

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < firstDataRow)
                continue;

            int columnCount = row.getPhysicalNumberOfCells();
            if (columnCount < maxColumnUsed)
                continue;

            try {
                String kitName = getKitName(row, kitColumn);
                String partNum = getPartNumber(row, partColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.isRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, poColumn);
                if (po.isEmpty())
                    continue;

                String grDate = getGrDateQty(row, grDateColumn);
                int grQty = getGrQty(row, grQtyColumn);

                int myQty = getMyQty(row, myQtyColumn);
                int applySet = getApplySet(row, setColumn);
                String remark = getRemark(row, remarkColumn);

                collection.addPurchase(kitName, partNum,
                        new PurchaseItem(rowNum, po, grDate, grQty, myQty, applySet, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close();
    }
}
