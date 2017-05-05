package lahome.rotateTool.Util;

import com.incesoft.tools.excel.xlsx.*;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.monitorjbl.xlsx.StreamingReader;
import javafx.collections.ObservableList;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

public class ExcelParser {
    private static final Logger log = LogManager.getLogger(ExcelParser.class.getName());

    private static final int sheetIdx = 0;

    private RotateCollection collection;

    private InputStream is;

    private String rotateFilePath;
    private int rotateFirstDataRow;
    private int rotateKitColumn;
    private int rotatePartColumn;
    private int rotatePmQtyColumn;
    private int rotateApQtyColumn;
    private int rotateRatioColumn;
    private int rotateApplySetColumn;
    private int rotateRemarkColumn;


    private String stockFilePath;
    private int stockFirstDataRow;
    private int stockKitColumn;
    private int stockPartColumn;
    private int stockPoColumn;
    private int stockStQtyColumn;
    private int stockLotColumn;
    private int stockDcColumn;
    private int stockApQtyColumn;
    private int stockRemarkColumn;


    private String purchaseFilePath;
    private int purchaseFirstDataRow;
    private int purchaseKitColumn;
    private int purchasePartColumn;
    private int purchasePoColumn;
    private int purchaseGrDateColumn;
    private int purchaseGrQtyColumn;
    private int purchaseApQtyColumn;
    private int purchaseApSetColumn;
    private int purchaseRemarkColumn;


    public ExcelParser(RotateCollection collection) {
        this.collection = collection;
    }

    public void setRotateConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                                String pmQtyColStr, String apQtyColStr, String ratioColStr, String applySetColStr,
                                String remarkColStr) {

        this.rotateFilePath = filePath;
        this.rotateFirstDataRow = firstDataRow - 1;

        rotateKitColumn = CellReference.convertColStringToIndex(kitColStr.toUpperCase());
        rotatePartColumn = CellReference.convertColStringToIndex(partColStr.toLowerCase());
        rotatePmQtyColumn = CellReference.convertColStringToIndex(pmQtyColStr.toUpperCase());
        rotateApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr.toUpperCase());
        rotateRatioColumn = CellReference.convertColStringToIndex(ratioColStr.toUpperCase());
        rotateApplySetColumn = CellReference.convertColStringToIndex(applySetColStr.toUpperCase());
        rotateRemarkColumn = CellReference.convertColStringToIndex(remarkColStr.toUpperCase());
    }

    public void setStockConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                               String poColStr, String stQtyColStr, String lotColStr, String dcColStr,
                               String apQtyColStr, String remarkColStr) {

        this.stockFilePath = filePath;
        this.stockFirstDataRow = firstDataRow - 1;

        stockKitColumn = CellReference.convertColStringToIndex(kitColStr);
        stockPartColumn = CellReference.convertColStringToIndex(partColStr);
        stockPoColumn = CellReference.convertColStringToIndex(poColStr);
        stockStQtyColumn = CellReference.convertColStringToIndex(stQtyColStr);
        stockLotColumn = CellReference.convertColStringToIndex(lotColStr);
        stockDcColumn = CellReference.convertColStringToIndex(dcColStr);
        stockApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr);
        stockRemarkColumn = CellReference.convertColStringToIndex(remarkColStr);
    }

    public void setPurchaseConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                                  String poColStr, String grDateColStr, String grQtyColStr,
                                  String apQtyColStr, String setColStr, String remarkColStr) {

        this.purchaseFilePath = filePath;
        this.purchaseFirstDataRow = firstDataRow - 1;

        purchaseKitColumn = CellReference.convertColStringToIndex(kitColStr);
        purchasePartColumn = CellReference.convertColStringToIndex(partColStr);
        purchasePoColumn = CellReference.convertColStringToIndex(poColStr);
        purchaseGrDateColumn = CellReference.convertColStringToIndex(grDateColStr);
        purchaseGrQtyColumn = CellReference.convertColStringToIndex(grQtyColStr);
        purchaseApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr);
        purchaseApSetColumn = CellReference.convertColStringToIndex(setColStr);
        purchaseRemarkColumn = CellReference.convertColStringToIndex(remarkColStr);
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

    private int getApQty(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
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
            return 0;
        }
        return (int) cellGetValue(cell);
    }

    private Workbook getWorkbook(String filePath) {

        File file = new File(filePath);
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

        Workbook wb = StreamingReader.builder()
                .rowCacheSize(100)     // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)

//        if (wb == null) {
//            log.warn(String.format("%s(%d) wb is null",
//                    Thread.currentThread().getStackTrace()[1].getMethodName(),
//                    Thread.currentThread().getStackTrace()[1].getLineNumber()));
//            return null;
//        }

        return wb;
    }

    private void close(Workbook wb) {

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
        try {
            if (cellType == Cell.CELL_TYPE_FORMULA) {
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        long value = (long) cell.getNumericCellValue();
                        return Long.toString(value);
                    case Cell.CELL_TYPE_STRING:
                        return cell.getStringCellValue();
                    case Cell.CELL_TYPE_FORMULA:
                        return cell.getStringCellValue();
                    default:
                        log.warn(String.format("unknown formula result type %d", cell.getCachedFormulaResultType()));
                        return "";
                }
            } else if (cellType == Cell.CELL_TYPE_STRING) {
                return cell.getStringCellValue();
            } else if (cellType == Cell.CELL_TYPE_NUMERIC) {
                long value = (long) cell.getNumericCellValue();
                return Long.toString(value);
            } else if (cellType == Cell.CELL_TYPE_BLANK) {
                return "";
            }
        } catch (Exception e) {
            log.error(String.format("%s:%s(%d) cell(%d,%d) is unknown type",
                    Thread.currentThread().getStackTrace()[1].getClassName(),
                    Thread.currentThread().getStackTrace()[1].getMethodName(),
                    Thread.currentThread().getStackTrace()[1].getLineNumber(),
                    cell.getRowIndex(), cell.getColumnIndex()));

            return "";
        }

        log.error(String.format("%s:%s(%d) cell(%d,%d) is unknown type",
                Thread.currentThread().getStackTrace()[1].getClassName(),
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


    public void loadRotateExcel() {

        Workbook wb = getWorkbook(rotateFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(sheetIdx);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < rotateFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();

            try {
                String kitName = getKitName(row, rotateKitColumn);
                String partNum = getPartNumber(row, rotatePartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                int pmQty = getExpectQty(row, rotatePmQtyColumn);
                if (pmQty == 0)
                    continue;

                String ratio = getRatio(row, rotateRatioColumn);
                int applySet = getApplySet(row, rotateApplySetColumn);
                String remark = getRemark(row, rotateRemarkColumn);

                collection.addRotate(new RotateItem(rowNum, kitName, partNum, pmQty, ratio, applySet, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close(wb);
    }

    public void loadStockExcel() {

        Workbook wb = getWorkbook(stockFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(sheetIdx);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }

        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < stockFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();

            try {
                String kitName = getKitName(row, stockKitColumn);
                String partNum = getPartNumber(row, stockPartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.belongToRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, stockPoColumn);
                if (po.isEmpty())
                    continue;

                int stockQty = getStockQty(row, stockStQtyColumn);
                String lot = getLot(row, stockLotColumn);
                String dc = getDc(row, stockDcColumn);
                int applyQty = getApQty(row, stockApQtyColumn);
                String remark = getRemark(row, stockRemarkColumn);

                collection.addStock(kitName, partNum,
                        new StockItem(rowNum, po, stockQty, lot, dc, applyQty, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close(wb);
    }

    public void loadPurchaseExcel() {

        Workbook wb = getWorkbook(purchaseFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(sheetIdx);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            if (rowNum < purchaseFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();
            try {
                String kitName = getKitName(row, purchaseKitColumn);
                String partNum = getPartNumber(row, purchasePartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.belongToRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, purchasePoColumn);
                if (po.isEmpty())
                    continue;

                String grDate = getGrDateQty(row, purchaseGrDateColumn);
                int grQty = getGrQty(row, purchaseGrQtyColumn);

                int applyQty = getApQty(row, purchaseApQtyColumn);
                int applySet = getApplySet(row, purchaseApSetColumn);
                String remark = getRemark(row, purchaseRemarkColumn);

                collection.addPurchase(kitName, partNum,
                        new PurchaseItem(rowNum, po, grDate, grQty, applyQty, applySet, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close(wb);
    }

    public void saveRotateExcel() {
        ObservableList<RotateItem> rotateItemList = collection.getRotateObsList();
        rotateItemList.sort(Comparator.comparing(RotateItem::getRowNum));

        try {
            OPCPackage pkg = OPCPackage.open(new File(rotateFilePath));
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            //Workbook wb = getWorkbook(rotateFilePath);

            Sheet sheet = wb.getSheetAt(sheetIdx);
            if (sheet == null) {
                log.error("sheet is null. Open sheet failed");
                return;
            }

            int itemIdx = 0;
            Cell cell;
            RotateItem rotateItem = rotateItemList.get(itemIdx);
            for (Row row : sheet) {
                int rowNum = row.getRowNum();
                if (rowNum == rotateItem.getRowNum()) {
                    cell = row.getCell(rotateApQtyColumn);
                    if (cell == null)
                        cell = row.createCell(rotateApQtyColumn, CellType.NUMERIC);

                    cell.setCellValue(rotateItem.getStockApplyQtyTotal());

                    itemIdx++;
                    if (itemIdx >= rotateItemList.size())
                        break;

                    rotateItem = rotateItemList.get(itemIdx);
                }

            }

            //FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
            //wb.write(fileOut);
            //fileOut.close();
            //pkg.flush();
            pkg.save(new File("D:\\Test.xlsx"));
            pkg.revert();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public static void testLoadALL(SimpleXLSXWorkbook workbook) {
        // medium data set,just load all at a time
        com.incesoft.tools.excel.xlsx.Sheet sheetToRead = workbook.getSheet(0);
        List<com.incesoft.tools.excel.xlsx.Cell[]> rows = sheetToRead.getRows();
        int rowPos = 0;
        for (com.incesoft.tools.excel.xlsx.Cell[] row : rows) {
            if (row != null) {
                printRow(rowPos, row);
            }
            rowPos++;
        }
    }

    public static Date cellGetDate(double numValue) {
        boolean time;

        // The number of days between 1 Jan 1900 and 1 March 1900. Excel thinks
        // the day before this was 29th Feb 1900, but it was 28th Feb 1900.
        // I guess the programmers thought nobody would notice that they
        // couldn't be bothered to program this dating anomaly properly
        final int nonLeapDay = 61;

        // The number of days between 01 Jan 1900 and 01 Jan 1970 - this gives
        // the UTC offset
        final int utcOffsetDays = 25569;

        // The number of milliseconds in  a day
        final long secondsInADay = 24 * 60 * 60;
        final long msInASecond = 1000;

        // This value represents the number of days since 01 Jan 1900
        time = Math.abs(numValue) < 1;

        // Work round a bug in excel.  Excel seems to think there is a date
        // called the 29th Feb, 1900 - but in actual fact this was not a leap year.
        // Therefore for values less than 61 in the 1900 date system,
        // add one to the numeric value
        if (!time && numValue < nonLeapDay) {
            numValue += 1;
        }

        // Convert this to the number of days since 01 Jan 1970
        double utcDays = numValue - utcOffsetDays;

        // Convert this into utc by multiplying by the number of milliseconds
        // in a day.  Use the round function prior to ms conversion due
        // to a rounding feature of Excel
        long utcValue = Math.round(utcDays * secondsInADay) * msInASecond;

        return new Date(utcValue);
    }

    public static void testIterateALL(SimpleXLSXWorkbook workbook) {
        // here we assume that the sheet contains too many rows which will leads
        // to memory overflow;
        // So we get sheet without loading all records
        com.incesoft.tools.excel.xlsx.Sheet sheetToRead = workbook.getSheet(0, false);
        com.incesoft.tools.excel.xlsx.Sheet.SheetRowReader reader = sheetToRead.newReader();
        com.incesoft.tools.excel.xlsx.Cell[] row;
        int rowPos = 0;
        while ((row = reader.readRow()) != null) {
            printRow(rowPos, row);
            rowPos++;
        }
    }

    private static void printRow(int rowPos, com.incesoft.tools.excel.xlsx.Cell[] row) {
        int cellPos = 0;
        for (com.incesoft.tools.excel.xlsx.Cell cell : row) {
            if (cell != null) {
                if (rowPos > 2 && cellPos == 5) {
                    System.out.println(com.incesoft.tools.excel.xlsx.Sheet.getCellId(rowPos, cellPos) + "="
                            + DateUtil.format(cellGetDate(Integer.valueOf(cell.getValue()))));
                }

                System.out.println(com.incesoft.tools.excel.xlsx.Sheet.getCellId(rowPos, cellPos) + "="
                        + cell.getValue());
            }
            cellPos++;
        }
    }


    public void savePurchaseExcel() {
        List<PurchaseItem> itemList = collection.getPurchaseItemList();
        itemList.sort(Comparator.comparing(PurchaseItem::getRowNum));

        try {
//            SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(new File("D:\\RotateFiles\\ZMMPO006_130307-131104.xlsx"));
//
//
//            com.incesoft.tools.excel.xlsx.Sheet sheet = workbook.getSheet(0, true);
//            List<com.incesoft.tools.excel.xlsx.Cell[]> rows = sheet.getRows();
//
//            FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
////            com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.Commiter commiter = workbook.newCommiter(fileOut);
////            commiter.beginCommit();
////            commiter.beginCommitSheet(sheet);
////            // merge it first,otherwise the modification will not take effect
////            commiter.commitSheetModifications();
//
//            int itemIdx = 0;
//            PurchaseItem item = itemList.get(itemIdx);
//            int rowNum = 0;
//            for (com.incesoft.tools.excel.xlsx.Cell[] row : rows) {
//                if (rowNum == item.getRowNum()) {
//                    sheet.modify(rowNum, rotateApQtyColumn, String.valueOf(item.getStockApplyQtyTotal()), null);
//
//                    itemIdx++;
//                    if (itemIdx >= itemList.size())
//                        break;
//
//                    item = itemList.get(itemIdx);
//                }
//                rowNum++;
//            }
//            workbook.commit(fileOut);
//            fileOut.close();
//
//        } catch (IOException e1) {
//            e1.printStackTrace();
//        } catch (XMLStreamException e1) {
//            e1.printStackTrace();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }

//
            OPCPackage pkg = OPCPackage.open(new File(purchaseFilePath));
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            Sheet sheet = wb.getSheetAt(sheetIdx);
            if (sheet == null) {
                log.error("sheet is null. Open sheet failed");
                return;
            }

            int itemIdx = 0;
            Cell cell;
            PurchaseItem item = itemList.get(itemIdx);
            for (Row row : sheet) {
                int rowNum = row.getRowNum();
                if (rowNum == item.getRowNum()) {
                    cell = row.getCell(purchaseApQtyColumn);
                    if (cell == null)
                        cell = row.createCell(purchaseApQtyColumn, CellType.NUMERIC);

                    cell.setCellValue(item.getApplyQty());

                    itemIdx++;
                    if (itemIdx >= itemList.size())
                        break;

                    item = itemList.get(itemIdx);
                }
            }

            //FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
            //wb.write(fileOut);
            //fileOut.close();
            //pkg.flush();
            pkg.save(new File("D:\\Test_ZMM.xlsx"));
            pkg.revert();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static List getOpenedWorkbooks() {
        List workbookNames = new ArrayList();
        try {
            ActiveXComponent xl = ActiveXComponent.connectToActiveInstance("Excel.Application");
            xl.setProperty("Visible", true);
            Dispatch xlo = xl.getObject();

            //xl.setProperty("Visible", new Variant(true));
            Dispatch workbooks = xl.getProperty("Workbooks").getDispatch();
            int openWorkBookCount = Dispatch.get(workbooks, "Count").getInt();
            if (openWorkBookCount == 0) {
                //throw new ExcelException("No open Excel workbooks");
            }
            for (int i = 1; i < openWorkBookCount; i++) {
                Dispatch oneworkbook = Dispatch.invoke(workbooks, "Item", Dispatch.Get, new Object[]{i}, new int[0]).getDispatch();
                workbookNames.add(Dispatch.get(oneworkbook, "Name").toString());
            }
            xlo.safeRelease();
        } catch (RuntimeException ex) {
            //throw new ExcelException("Unable to read from Excel Process, trye again later.", ex);
        }

        return workbookNames;
    }

    public String excelColumnToLetter(int col) {

        int excelColNum = col + 1;
        StringBuilder colRef = new StringBuilder(2);
        int colRemain = excelColNum;

        while(colRemain > 0) {
            int thisPart = colRemain % 26;
            if(thisPart == 0) {
                thisPart = 26;
            }

            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char)(thisPart + 64);
            colRef.insert(0, colChar);
        }

        return colRef.toString();
    }

    public void writeCell(Dispatch sheet, int row, int col, String data) {
        String pos = excelColumnToLetter(col) + row;
        Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{pos}, new int[1]).toDispatch();
        Dispatch.put(cell, "Value", data);
    }

    public void test() {
        List<PurchaseItem> itemList = collection.getPurchaseItemList();
        itemList.sort(Comparator.comparing(PurchaseItem::getRowNum));


        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant("D:\\RotateFiles\\ZMMPO006_130307-131104.xlsx")).toDispatch();
            getOpenedWorkbooks();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{1}, new int[1]).toDispatch();
            String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            Dispatch a1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{"A1"}, new int[1]).toDispatch();
            Dispatch.put(a1, "Value", "123.456");

            String a = CellReference.convertNumToColString(27);
            String b = excelColumnToLetter(27);

            int itemIdx = 0;
            PurchaseItem item = itemList.get(itemIdx);
            for (int rowNum = 1; rowNum < rowCount; rowNum++) {
                if (rowNum == item.getRowNum()) {
                    writeCell(sheet, rowNum, purchaseApQtyColumn, String.valueOf(item.getApplyQty()));

                    itemIdx++;
                    if (itemIdx >= itemList.size())
                        break;

                    item = itemList.get(itemIdx);
                }
            }
            // Saves and closes
//            Dispatch.call(workBook, "Save");
//            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
//            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }

    }
}
