package lahome.rotateTool.Util.Excel;

import com.incesoft.tools.excel.xlsx.*;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.monitorjbl.xlsx.StreamingReader;
import javafx.beans.property.DoubleProperty;
import javafx.beans.property.SimpleDoubleProperty;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelParser {
    private static final Logger log = LogManager.getLogger(ExcelParser.class.getName());


    private RotateCollection collection;

    private InputStream is;

    public ExcelParser(RotateCollection collection) {
        this.collection = collection;
    }

    public static int convertColStringToIndex(String ref) {
        int retValue = 0;
        char[] refs = ref.toUpperCase(Locale.ROOT).toCharArray();

        for(int k = 0; k < refs.length; ++k) {
            char ch = refs[k];
            if(ch == 36) {
                if(k != 0) {
                    throw new IllegalArgumentException("Bad col ref format '" + ref + "'");
                }
            } else {
                retValue = retValue * 26 + ch - 65 + 1;
            }
        }

        return retValue - 1;
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
            return lahome.rotateTool.Util.DateUtil.format(date);
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

    public void loadRotateExcel(DoubleProperty rotateProgressProperty) {

        Workbook wb = getWorkbook(ExcelSettings.rotateFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(ExcelSettings.rotateSheetIndex - 1);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }

        int rowCount = sheet.getLastRowNum();
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            rotateProgressProperty.set(Math.min(1, (double)rowNum / rowCount));

            if (rowNum < ExcelSettings.rotateFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();

            try {
                String kitName = getKitName(row, ExcelSettings.rotateKitColumn);
                String partNum = getPartNumber(row, ExcelSettings.rotatePartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                int pmQty = getExpectQty(row, ExcelSettings.rotatePmQtyColumn);
                if (pmQty == 0)
                    continue;

                String ratio = getRatio(row, ExcelSettings.rotateRatioColumn);
                String remark = getRemark(row, ExcelSettings.rotateRemarkColumn);

                collection.addRotate(new RotateItem(rowNum, kitName, partNum, pmQty, ratio, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        rotateProgressProperty.set(1.0);
        close(wb);
    }

    public void loadStockExcel(DoubleProperty stockProgressProperty) {

        Workbook wb = getWorkbook(ExcelSettings.stockFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(ExcelSettings.stockSheetIndex - 1);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }

        int rowCount = sheet.getLastRowNum();
        for (Row row : sheet) {
            int rowNum = row.getRowNum();
            stockProgressProperty.set((double)rowNum / rowCount);

            if (rowNum < ExcelSettings.stockFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();

            try {
                String kitName = getKitName(row, ExcelSettings.stockKitColumn);
                String partNum = getPartNumber(row, ExcelSettings.stockPartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.belongToRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, ExcelSettings.stockPoColumn);
                if (po.isEmpty())
                    continue;

                int stockQty = getStockQty(row, ExcelSettings.stockStQtyColumn);
                String lot = getLot(row, ExcelSettings.stockLotColumn);
                String dc = getDc(row, ExcelSettings.stockDcColumn);
                int applyQty = getApQty(row, ExcelSettings.stockApQtyColumn);
                String remark = getRemark(row, ExcelSettings.stockRemarkColumn);

                collection.addStock(kitName, partNum,
                        new StockItem(rowNum, po, stockQty, lot, dc, applyQty, remark));

            } catch (Exception e) {
                log.error("failed!", e);
            }
        }

        close(wb);
    }

    public void loadPurchaseExcel(DoubleProperty purchaseProgressProperty) {

        Workbook wb = getWorkbook(ExcelSettings.purchaseFilePath);
        if (wb == null) {
            log.error("workbook is null. Open sheet failed");
            return;
        }

        Sheet sheet = wb.getSheetAt(ExcelSettings.purchaseSheetIndex - 1);
        if (sheet == null) {
            log.error("sheet is null. Open sheet failed");
            return;
        }

        int rowCount = sheet.getLastRowNum();
        for (Row row : sheet) {

            int rowNum = row.getRowNum();
            purchaseProgressProperty.set((double)rowNum / rowCount);

            if (rowNum < ExcelSettings.purchaseFirstDataRow)
                continue;

            //int columnCount = row.getPhysicalNumberOfCells();
            try {
                String kitName = getKitName(row, ExcelSettings.purchaseKitColumn);
                String partNum = getPartNumber(row, ExcelSettings.purchasePartColumn);
                if (kitName.isEmpty() && partNum.isEmpty())
                    continue;

                if (!collection.belongToRotateItem(kitName, partNum))
                    continue;

                String po = getPo(row, ExcelSettings.purchasePoColumn);
                if (po.isEmpty())
                    continue;

                String grDate = getGrDateQty(row, ExcelSettings.purchaseGrDateColumn);
                int grQty = getGrQty(row, ExcelSettings.purchaseGrQtyColumn);

                int applyQty = getApQty(row, ExcelSettings.purchaseApQtyColumn);
                String remark = getRemark(row, ExcelSettings.purchaseRemarkColumn);

                collection.addPurchase(kitName, partNum,
                        new PurchaseItem(rowNum, po, grDate, grQty, applyQty, remark));

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
            OPCPackage pkg = OPCPackage.open(new File(ExcelSettings.rotateFilePath));
            XSSFWorkbook wb = new XSSFWorkbook(pkg);
            //Workbook wb = getWorkbook(rotateFilePath);

            Sheet sheet = wb.getSheetAt(ExcelSettings.rotateSheetIndex - 1);
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
                    cell = row.getCell(ExcelSettings.rotateApQtyColumn);
                    if (cell == null)
                        cell = row.createCell(ExcelSettings.rotateApQtyColumn, CellType.NUMERIC);

                    cell.setCellValue(rotateItem.getStockApplyQtyTotal());

                    itemIdx++;
                    if (itemIdx >= rotateItemList.size())
                        break;

                    rotateItem = rotateItemList.get(itemIdx);
                }

            }

            //FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
            //wb.write(fileOut);
            //fileOut.clear();
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
                            + lahome.rotateTool.Util.DateUtil.format(cellGetDate(Integer.valueOf(cell.getValue()))));
                }

                System.out.println(com.incesoft.tools.excel.xlsx.Sheet.getCellId(rowPos, cellPos) + "="
                        + cell.getValue());
            }
            cellPos++;
        }
    }


//    public void savePurchaseExcel() {
//        List<PurchaseItem> itemList = collection.getPurchaseItemList();
//        itemList.sort(Comparator.comparing(PurchaseItem::getRowNum));
//
//        try {
//            SimpleXLSXWorkbook workbook = new SimpleXLSXWorkbook(new File("D:\\RotateFiles\\ZMMPO006_130307-131104.xlsx"));
//
//
//            com.incesoft.tools.excel.xlsx.Sheet sheet = workbook.getSheet(0, true);
//            List<com.incesoft.tools.excel.xlsx.Cell[]> rows = sheet.getRows();
//
//            FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
//            com.incesoft.tools.excel.xlsx.SimpleXLSXWorkbook.Commiter commiter = workbook.newCommiter(fileOut);
//            commiter.beginCommit();
//            commiter.beginCommitSheet(sheet);
//            // merge it first,otherwise the modification will not take effect
//            commiter.commitSheetModifications();

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
//            fileOut.clear();

//        } catch (IOException e1) {
//            e1.printStackTrace();
//        } catch (XMLStreamException e1) {
//            e1.printStackTrace();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }

//
//            OPCPackage pkg = OPCPackage.open(new File(ExcelSettings.purchaseFilePath));
//            XSSFWorkbook wb = new XSSFWorkbook(pkg);
//            Sheet sheet = wb.getSheetAt(ExcelSettings.purchaseSheetIndex - 1);
//            if (sheet == null) {
//                log.error("sheet is null. Open sheet failed");
//                return;
//            }
//
//            int itemIdx = 0;
//            Cell cell;
//            PurchaseItem item = itemList.get(itemIdx);
//            for (Row row : sheet) {
//                int rowNum = row.getRowNum();
//                if (rowNum == item.getRowNum()) {
//                    cell = row.getCell(ExcelSettings.purchaseApQtyColumn);
//                    if (cell == null)
//                        cell = row.createCell(ExcelSettings.purchaseApQtyColumn, CellType.NUMERIC);
//
//                    cell.setCellValue(item.getApplyQty());
//
//                    itemIdx++;
//                    if (itemIdx >= itemList.size())
//                        break;
//
//                    item = itemList.get(itemIdx);
//                }
//            }
//
//            //FileOutputStream fileOut = new FileOutputStream("D:\\AWG.xlsx");
//            //wb.write(fileOut);
//            //fileOut.clear();
//            //pkg.flush();
//            pkg.save(new File("D:\\Test_ZMM.xlsx"));
//            pkg.revert();
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }

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





}
