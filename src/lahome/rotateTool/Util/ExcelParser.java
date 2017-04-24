package lahome.rotateTool.Util;

import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.*;

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
            is = new FileInputStream(file.getPath());
        } catch (FileNotFoundException e) {
            log.error(String.format("File %s opened failed", file.getPath()));
            return null;
        }

        Workbook wb;
        try {
            wb = WorkbookFactory.create(is);
        } catch (IOException | InvalidFormatException e) {
            log.error("failed!", e);
            return null;
        }
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
            return cell.getStringCellValue();
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


    public void loadRotateExcel(File file, int firstDataRow, String kitColStr, String partColStr, String pmQty) {

        firstDataRow = firstDataRow - 1;
        int kitColumn = CellReference.convertColStringToIndex(kitColStr.toUpperCase());
        int partColumn = CellReference.convertColStringToIndex(partColStr.toLowerCase());
        int pmQtyColumn = CellReference.convertColStringToIndex(pmQty.toUpperCase());

        int maxColumnUsed = Math.max(Math.max(kitColumn, partColumn), pmQtyColumn);

        Sheet sheet = getSheet(file);
        if (sheet == null) {
            log.error("open sheet failed");
            return;
        }

        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = firstDataRow; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            int rowNum = row.getRowNum();
            log.debug(String.format("Rotate row %d", rowNum));

            int columnCount = row.getPhysicalNumberOfCells();
            if (columnCount < maxColumnUsed)
                continue;

            try {
                String kit = getKitName(row, kitColumn);
                String part = getPartNumber(row, partColumn);
                if (kit.isEmpty() && part.isEmpty())
                    continue;

                int qty = getExpectQty(row, pmQtyColumn);

                if (qty == 0)
                    continue;

                RotateItem item = new RotateItem(row, kit, part, qty);
                collection.addRotate(item);

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

        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = firstDataRow; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            int rowNum = row.getRowNum();
            log.debug(String.format("Stock row %d", rowNum));

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

                collection.addStock(row, kitName, partNumber, po, stockQty, dc, myQty);

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

        int rowCount = sheet.getPhysicalNumberOfRows();
        for (int i = firstDataRow; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null)
                continue;

            int rowNum = row.getRowNum();
            log.debug(String.format("Purchase row %d", rowNum));

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


                collection.addPurchase(row, kit, part, po, date, grQty);

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }
    }
}
