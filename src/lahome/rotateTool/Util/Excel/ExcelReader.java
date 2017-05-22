package lahome.rotateTool.Util.Excel;

import javafx.beans.property.DoubleProperty;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.xml.sax.SAXException;

import java.util.Date;
import java.util.Locale;

@SuppressWarnings("unused")
public class ExcelReader {
    private static final Logger log = LogManager.getLogger(ExcelReader.class.getName());

    private int lastRow = 0;
    private double oldProgress;

    private RotateCollection collection;

    public ExcelReader(RotateCollection collection) {
        this.collection = collection;
    }

    public static int convertColStringToIndex(String ref) {
        int retValue = 0;
        char[] refs = ref.toUpperCase(Locale.ROOT).toCharArray();

        for (int k = 0; k < refs.length; ++k) {
            char ch = refs[k];
            if (ch == 36) {
                if (k != 0) {
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

        return kitName.trim().toUpperCase();
    }

    private String getPartNumber(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return "";
        }

        return cellGetString(cell).trim().toUpperCase();
    }

    private int getBacklog(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return -1;
        }

        return (int) cellGetValue(cell);
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

        return cellGetString(cell).trim().toUpperCase();
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

    private int getDc(Row row, int column) throws Exception {
        Cell cell = row.getCell(column);
        if (cell == null) {
            return 0;
        }
        int dc = (int) cellGetValue(cell);
        if (dc / 100 > 99)
            return 0;
        if (dc / 100 > 53)
            return 0;

        return dc;
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

    @SuppressWarnings("deprecation")
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

    @SuppressWarnings("deprecation")
    private double cellGetValue(Cell cell) {
        try {
            int cellType = cell.getCellType();
            if (cellType == Cell.CELL_TYPE_FORMULA) {
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        return cell.getNumericCellValue();
                    case Cell.CELL_TYPE_STRING:
                        return Double.valueOf(numTrim(cell.getStringCellValue().trim()));
                    default:
                        return 0;
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
        } catch (NumberFormatException e) {
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
            if (!Character.isDigit(sb.charAt(i)))
                sb.deleteCharAt(i);
            else
                i++;
        }
        return sb.toString();
    }

//    private boolean isNumber(char c) {
//        return (c >= '0' && c <= '9');
//    }

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

    public void loadRotateExcel(DoubleProperty rotateProgress) {
        lastRow = 0;
        oldProgress = 0;

        StreamingReader reader = new StreamingReader() {
            @Override
            public void getRows(int sheetIndex, Row row) {
                int rowNum = row.getRowNum() + 1;

                if (lastRow > 0) {
                    double progress = (double) rowNum / lastRow;
                    if (progress >= oldProgress + 0.01) {
                        rotateProgress.set(progress);
                        oldProgress = progress;
                    }
                }

                if (rowNum < ExcelSettings.rotateFirstDataRow)
                    return;

                // System.out.println("Sheet:" + sheetIndex + ", Row:" + rowNum + ", Data:" + row);
                try {
                    String kitName = getKitName(row, ExcelSettings.rotateKitColumn);
                    String partNum = getPartNumber(row, ExcelSettings.rotatePartColumn);
                    if (kitName.isEmpty() && partNum.isEmpty())
                        return;

                    int pmQty = getExpectQty(row, ExcelSettings.rotatePmQtyColumn);
                    if (pmQty == 0)
                        return;

                    int backlog = getBacklog(row, ExcelSettings.rotateBacklogColumn);

                    String ratio = getRatio(row, ExcelSettings.rotateRatioColumn);
                    String remark = getRemark(row, ExcelSettings.rotateRemarkColumn);

                    collection.addRotate(new RotateItem(rowNum, kitName, partNum, backlog, pmQty, ratio, remark));

                } catch (Exception e) {
                    log.error("failed!", e);
                }
            }

            @Override
            public void getLastRowNum(int lastRowNum) {
                lastRow = lastRowNum;
            }

            @Override
            public void endDocument() throws SAXException {
                super.endDocument();
                rotateProgress.set(1.0);
            }
        };

        try {
            reader.processOneSheet(ExcelSettings.rotateFilePath, ExcelSettings.rotateSheetIndex);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void loadStockExcel(DoubleProperty stockProgress) {
        lastRow = 0;
        oldProgress = 0;

        StreamingReader reader = new StreamingReader() {
            @Override
            public void getRows(int sheetIndex, Row row) {
                int rowNum = row.getRowNum() + 1;

                if (lastRow > 0) {
                    double progress = (double) rowNum / lastRow;
                    if (progress >= oldProgress + 0.01) {
                        stockProgress.set(progress);
                        oldProgress = progress;
                    }
                }

                if (rowNum < ExcelSettings.stockFirstDataRow)
                    return;

                // System.out.println("Sheet:" + sheetIndex + ", Row:" + rowNum + ", Data:" + row);
                try {
                    String kitName = getKitName(row, ExcelSettings.stockKitColumn);
                    String partNum = getPartNumber(row, ExcelSettings.stockPartColumn);
                    if (kitName.isEmpty() && partNum.isEmpty())
                        return;

                    RotateItem rotateItem = collection.getRotateItem(kitName, partNum);
                    if (rotateItem == null)
                        return;

                    String po = getPo(row, ExcelSettings.stockPoColumn);
                    if (po.isEmpty())
                        return;

                    int stockQty = getStockQty(row, ExcelSettings.stockStQtyColumn);
                    String lot = getLot(row, ExcelSettings.stockLotColumn);
                    int dc = getDc(row, ExcelSettings.stockDcColumn);
                    int applyQty = getApQty(row, ExcelSettings.stockApQtyColumn);
                    String remark = getRemark(row, ExcelSettings.stockRemarkColumn);

                    collection.addStock(rotateItem,
                            new StockItem(rowNum, po, stockQty, lot, dc, applyQty, remark));
                } catch (Exception e) {

                    log.error("failed!", e);
                }
            }

            @Override
            public void getLastRowNum(int lastRowNum) {
                lastRow = lastRowNum;
            }

            @Override
            public void endDocument() throws SAXException {
                super.endDocument();
                stockProgress.set(1.0);
            }
        };

        try {
            reader.processOneSheet(ExcelSettings.stockFilePath, ExcelSettings.stockSheetIndex);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public void loadPurchaseExcel(DoubleProperty purchaseProgress) {
        lastRow = 0;
        oldProgress = 0;

        StreamingReader reader = new StreamingReader() {
            @Override
            public void getRows(int sheetIndex, Row row) {
                int rowNum = row.getRowNum() + 1;

                if (lastRow > 0) {
                    double progress = (double) rowNum / lastRow;
                    if (progress >= oldProgress + 0.01) {
                        purchaseProgress.set(progress);
                        oldProgress = progress;
                    }
                }

                if (rowNum < ExcelSettings.rotateFirstDataRow)
                    return;

                // System.out.println("Sheet:" + sheetIndex + ", Row:" + rowNum + ", Data:" +rowList);
                try {
                    String kitName = getKitName(row, ExcelSettings.purchaseKitColumn);
                    String partNum = getPartNumber(row, ExcelSettings.purchasePartColumn);
                    if (kitName.isEmpty() && partNum.isEmpty())
                        return;

                    RotateItem rotateItem = collection.getRotateItem(kitName, partNum);
                    if (rotateItem == null)
                        return;

                    String po = getPo(row, ExcelSettings.purchasePoColumn);
                    if (po.isEmpty())
                        return;

                    String grDate = getGrDateQty(row, ExcelSettings.purchaseGrDateColumn);
                    int grQty = getGrQty(row, ExcelSettings.purchaseGrQtyColumn);

                    int applyQty = getApQty(row, ExcelSettings.purchaseApQtyColumn);
                    String remark = getRemark(row, ExcelSettings.purchaseRemarkColumn);

                    collection.addPurchase(rotateItem,
                            new PurchaseItem(rowNum, po, grDate, grQty, applyQty, remark));
                } catch (Exception e) {
                    log.error("failed!", e);
                }
            }

            @Override
            public void getLastRowNum(int lastRowNum) {
                lastRow = lastRowNum;
            }

            @Override
            public void endDocument() throws SAXException {
                super.endDocument();
                purchaseProgress.set(1.0);
            }
        };

        try {
            reader.processOneSheet(ExcelSettings.purchaseFilePath, ExcelSettings.purchaseSheetIndex);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
