package lahome.rotateTool.Util.Excel;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import javafx.beans.property.DoubleProperty;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

@SuppressWarnings("WeakerAccess")
public class ExcelSaver {
    private static final Logger log = LogManager.getLogger(PurchaseItem.class.getName());

    private RotateCollection collection;
    private Dispatch currentSheet;

    public ExcelSaver(RotateCollection collection) {
        this.collection = collection;
    }

    public static String convertNumToColString(int col) {

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

    public void writeCell(int row, int col, String data) {
        try {
            String pos = convertNumToColString(col) + row;
            Dispatch cell = Dispatch.invoke(currentSheet, "Range", Dispatch.Get, new Object[]{pos}, new int[1]).toDispatch();
            Dispatch.put(cell, "Value", data);
        } catch (Exception e) {
            log.error("failed!", e);
        }
    }

    public void saveRotateToExcel(DoubleProperty rotateProgressProperty) {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelSettings.rotateFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            //int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelSettings.rotateSheetIndex}, new int[1]).toDispatch();
            //String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            //Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            //Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            //int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            int itemCount = 0;
            int itemsTotal = collection.getRotateObsList().size();
            for (RotateItem item : collection.getRotateObsList()) {
                rotateProgressProperty.set((double)++itemCount / itemsTotal);

                if (item.isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelSettings.rotateApQtyColumn, String.valueOf(item.getStockApplyQtyTotal()));
                writeCell(rowNum, ExcelSettings.rotateRemarkColumn, item.getRemark());
                if (item.isKit()) {
                    writeCell(rowNum, ExcelSettings.rotateApplySetColumn, String.valueOf(item.getApplySet()));
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
            ComThread.quitMainSTA();
            ComThread.Release();
            rotateProgressProperty.set(1.0);
        }
    }

    public void saveStockToExcel(DoubleProperty stockProgressProperty) {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelSettings.stockFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            //int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelSettings.stockSheetIndex}, new int[1]).toDispatch();
            //String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            //Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            //Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            //int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            int itemCount = 0;
            int itemsTotal = collection.getStockItemList().size();
            for (StockItem item : collection.getStockItemList()) {
                stockProgressProperty.set((double)++itemCount / itemsTotal);
                if (item.getRotateItem().isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelSettings.stockApQtyColumn, String.valueOf(item.getApplyQty()));
                writeCell(rowNum, ExcelSettings.stockRemarkColumn, item.getRemark());
            }

            // Saves and closes
//            Dispatch.call(workBook, "Save");
//            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
//            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //excel.invoke("Quit", new Variant[0]);
            ComThread.quitMainSTA();
            ComThread.Release();
            stockProgressProperty.set(1.0);
        }
    }

    public void savePurchaseToExcel(DoubleProperty purchaseProgressProperty) {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelSettings.purchaseFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            //int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelSettings.purchaseSheetIndex}, new int[1]).toDispatch();
            //String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            //Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            //Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            //int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            int itemCount = 0;
            int itemsTotal = collection.getPurchaseItemList().size();
            for (PurchaseItem item : collection.getPurchaseItemList()) {
                purchaseProgressProperty.set((double)++itemCount / itemsTotal);

                if (item.getRotateItem().isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelSettings.purchaseApQtyColumn, String.valueOf(item.getApplyQty()));
                writeCell(rowNum, ExcelSettings.purchaseRemarkColumn, item.getRemark());

                if (item.getRotateItem().isKit()) {
                    writeCell(rowNum, ExcelSettings.purchaseApSetColumn, String.valueOf(item.getApplySet()));
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
            ComThread.quitMainSTA();
            ComThread.Release();
            purchaseProgressProperty.set(1.0);
        }

    }
}
