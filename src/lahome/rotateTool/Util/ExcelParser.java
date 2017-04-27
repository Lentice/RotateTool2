package lahome.rotateTool.Util;

import com.monitorjbl.xlsx.StreamingReader;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Date;

public class ExcelParser {
    private static final Logger log = LogManager.getLogger(ExcelParser.class.getName());

    private static final int sheetIdx = 0;

    private RotateCollection collection;

    public ExcelParser(RotateCollection collection) {
        this.collection = collection;
    }

    private String getKitName(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("kit is null at row %d", row.getRowNum()));
        }

        String kitName = cellGetString(cell).trim();
        if (kitName.contains("空白"))
            kitName = "";

        return kitName;
    }

    private String getPartNumber(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("part is null at row %d", row.getRowNum()));
        }

        return cellGetString(cell).trim();
    }

    private int getExpectQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("qty is null at row %d", row.getRowNum()));
        }
        return (int) cellGetValue(cell);
    }

    private String getPo(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("po is null at row %d", row.getRowNum()));
        }

        return cellGetString(cell).trim();
    }

    private int getStockQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("stockQty is null at row %d", row.getRowNum()));
        }
        return (int) cellGetValue(cell);
    }

    private int getDc(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("dc is null at row %d", row.getRowNum()));
        }
        return (int) cellGetValue(cell);
    }

    private int getMyQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            //throw new Exception(String.format("myQty is null at row %d", row.getRowNum()));
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private String getGrDateQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            throw new Exception(String.format("stockQty is null at row %d", row.getRowNum()));
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

        InputStream is;
        try {
            is = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            log.error(String.format("File %s opened failed", file.getPath()));
            return null;
        }

        Workbook wb;
        wb = StreamingReader.builder()
                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)
        if (wb == null) {
            log.warn(String.format("%s(%d) wb is null",
                    Thread.currentThread().getStackTrace()[1].getMethodName(),
                    Thread.currentThread().getStackTrace()[1].getLineNumber()));
            return null;
        }

        return wb.getSheetAt(sheetIdx);
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


    public void loadRotateExcel(File file, int firstDataRow, String kitColStr, String partColStr, String pmQtyColStr) {

        firstDataRow = firstDataRow - 1;
        int kitColumn = CellReference.convertColStringToIndex(kitColStr.toUpperCase());
        int partColumn = CellReference.convertColStringToIndex(partColStr.toLowerCase());
        int pmQtyColumn = CellReference.convertColStringToIndex(pmQtyColStr.toUpperCase());

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
                String kit = getKitName(row, kitColumn);
                String part = getPartNumber(row, partColumn);
                if (kit.isEmpty() && part.isEmpty())
                    continue;

                int pmQty = getExpectQty(row, pmQtyColumn);

                if (pmQty == 0)
                    continue;

                collection.addRotate(rowNum, kit, part, pmQty);

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }
    }

    public void loadStockExcel(File file, int firstDataRow, String kitColStr, String partColStr,
                               String poColStr, String stockQtyColStr, String dcColStr, String myQtyColStr) {
        firstDataRow = firstDataRow - 1;

        int kitColumn = CellReference.convertColStringToIndex(kitColStr);
        int partColumn = CellReference.convertColStringToIndex(partColStr);
        int poColumn = CellReference.convertColStringToIndex(poColStr);
        int stockQtyColumn = CellReference.convertColStringToIndex(stockQtyColStr);
        int dcColumn = CellReference.convertColStringToIndex(dcColStr);
        int myQtyColumn = CellReference.convertColStringToIndex(myQtyColStr);

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
                String partNumber = getPartNumber(row, partColumn);
                if (kitName.isEmpty() && partNumber.isEmpty())
                    continue;

                if (!collection.isRotateItem(kitName, partNumber))
                    continue;

                String po = getPo(row, poColumn);
                if (po.isEmpty())
                    continue;

                int stockQty = getStockQty(row, stockQtyColumn);
                int dc = getDc(row, dcColumn);
                int myQty = getMyQty(row, myQtyColumn);

                collection.addStock(rowNum, kitName, partNumber, po, stockQty, dc, myQty);

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }
    }

    public void loadPurchaseExcel(File file, int firstDataRow, String kitColStr, String partColStr,
                                  String poColStr, String grDateColStr, String grQtyColStr) {
        firstDataRow = firstDataRow - 1;
        int kitColumn = CellReference.convertColStringToIndex(kitColStr);
        int partColumn = CellReference.convertColStringToIndex(partColStr);
        int poColumn = CellReference.convertColStringToIndex(poColStr);
        int grDateColumn = CellReference.convertColStringToIndex(grDateColStr);
        int grQtyColumn = CellReference.convertColStringToIndex(grQtyColStr);

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
                String kit = getKitName(row, kitColumn);
                String part = getPartNumber(row, partColumn);
                if (kit.isEmpty() && part.isEmpty())
                    continue;

                if (!collection.isRotateItem(kit, part))
                    continue;

                String po = getPo(row, poColumn);
                if (po.isEmpty())
                    continue;

                String date = getGrDateQty(row, grDateColumn);
                int grQty = getGrQty(row, grQtyColumn);


                collection.addPurchase(rowNum, kit, part, po, date, grQty);

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }
    }
}
