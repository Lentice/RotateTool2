package lahome.rotateTool.Util.Excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.Iterator;

public abstract class StreamingReader extends DefaultHandler {

    // 共享字符串表
    private SharedStringsTable sst;
    private StylesTable stylesTable;

    private final DataFormatter dataFormatter = new DataFormatter();

    private boolean useDate1904 = false;

    private String lastContents;

    private int sheetIndex = -1;
    private int curRow = 0;
    private int curCol = 0;
    private String cellStyleString;
    private StreamingCell currentCell;
    private StreamingRow row;

    public void processAllSheets(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        StylesTable styles = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst, styles);
        Iterator<InputStream> sheets = r.getSheetsData();
        while (sheets.hasNext()) {
            curRow = 0;
            sheetIndex++;
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            parser.parse(sheetSource);
            sheet.close();
        }

        pkg.revert();
    }

    public void processOneSheet(String filename, int sheetIndex) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        StylesTable styles = r.getStylesTable();
        SharedStringsTable sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser(sst, styles);

        // To look up the Sheet Name / Sheet Order / rID,
        //  you need to processOneSheet the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet = r.getSheet("rId" + sheetIndex);
        this.sheetIndex++;
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);

        sheet.close();
        pkg.revert();
    }


    public XMLReader fetchSheetParser(SharedStringsTable sst, StylesTable stylesTable) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        this.sst = sst;
        this.stylesTable = stylesTable;
        parser.setContentHandler(this);
        return parser;
    }

    private void setFormatString(lahome.rotateTool.Util.Excel.StreamingCell cell) {
        XSSFCellStyle style = null;

        if (cellStyleString != null) {
            style = stylesTable.getStyleAt(Integer.parseInt(cellStyleString));
        } else if (stylesTable.getNumCellStyles() > 0) {
            style = stylesTable.getStyleAt(0);
        }

        if (style != null) {
            cell.setNumericFormatIndex(style.getDataFormat());
            String formatString = style.getDataFormatString();

            if (formatString != null) {
                cell.setNumericFormat(formatString);
            } else {
                cell.setNumericFormat(BuiltinFormats.getBuiltinFormat(cell.getNumericFormatIndex()));
            }
        } else {
            cell.setNumericFormatIndex(null);
            cell.setNumericFormat(null);
        }
    }

    private String formattedContents() {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                int idx = Integer.parseInt(lastContents);
                return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(lastContents).toString();
            case "str":         //forumla type
                return '"' + lastContents + '"';
            case "e":           //error type
                return "ERROR:  " + lastContents;
            case "n":           //numeric type
//                if(currentCell.getNumericFormat() != null && lastContents.length() > 0) {
//                    return dataFormatter.formatRawCellContents(
//                            Double.parseDouble(lastContents),
//                            currentCell.getNumericFormatIndex(),
//                            currentCell.getNumericFormat());
//                } else {
                return lastContents;
//                }
            default:
                return lastContents;
        }
    }

    private String unformattedContents() {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                int idx = Integer.parseInt(lastContents);
                return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(lastContents).toString();
            default:
                return lastContents;
        }
    }

    private void getCoordinate(String coord) {
        for (int i = 1; i < coord.length(); i++) {
            if (isNumber(coord.charAt(i))) {
                try {
                    if (curCol == 0)
                        curRow = Integer.valueOf(coord.substring(i)) - 1;

                    curCol = CellReference.convertColStringToIndex(coord.substring(0, i));
                } catch (NumberFormatException ignore) {
                }
                break;
            }
        }
    }

    private boolean isNumber(char c) {
        return (c >= '0' && c <= '9');
    }

    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        // c => cell
        if ("c".equals(name)) {

            String postion = attributes.getValue("r");
            if (postion != null) {
                getCoordinate(postion);
            }
            currentCell = new StreamingCell(curCol, curRow, useDate1904);
            //setFormatString(currentCell);

            String cellType = attributes.getValue("t");
            if (cellType != null) {
                currentCell.setType(cellType);
            } else {
                currentCell.setType("n");
            }

            //cellStyleString = attributes.getValue("s");
            //if (cellStyleString != null) {
                //try {
                    //int index = Integer.parseInt(cellStyleString);
                    //currentCell.setCellStyle(stylesTable.getStyleAt(index));
                //} catch (NumberFormatException nfe) {
                    //log.warn("Ignoring invalid style index {}", indexStr);
                //}
            //}
        } else if ("dimension".equals(name)) {
            String ref = attributes.getValue("ref");
            // ref is formatted as A1 or A1:F25. Take the last numbers of this string and use it as lastRowNum
            for (int i = ref.length() - 1; i >= 0; i--) {
                if (!Character.isDigit(ref.charAt(i))) {
                    try {
                        int lastRowNum = Integer.parseInt(ref.substring(i + 1)) - 1;
                        getLastRowNum(lastRowNum);
                    } catch (NumberFormatException ignore) {
                    }
                    break;
                }
            }
        } else if ("row".equals(name)) {
            String rowNum = attributes.getValue("r");
            try {
                curRow = Integer.parseInt(rowNum) - 1;
            } catch (NumberFormatException ignore) {
            }
        } else if ("f".equals(name)) {  // formula
            currentCell.setType("str");
        }

        String date1904 = attributes.getValue("useDate1904");
        if (date1904 != null && "1".equals(date1904)) {
            useDate1904 = true;
        }

        // clear
        lastContents = "";
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {

        if ("v".equals(name) || "t".equals(name)) {

            currentCell.setRawContents(unformattedContents());
            currentCell.setContents(formattedContents());

        } else if ("row".equals(name)) {
            getRows(sheetIndex, row);
            row = null;
            curCol = 0;
            curRow++;
        } else if ("c".equals(name)) {
            if (row == null) {
                row = new StreamingRow(curRow);
            }
            row.addCell(curCol, currentCell);
            curCol++;
        } else if ("f".equals(name)) {
            currentCell.setFormula(lastContents);
        }
    }

    public void characters(char[] ch, int start, int length)
            throws SAXException {
        lastContents += new String(ch, start, length);
    }

    @Override
    public void endDocument() throws SAXException {
        super.endDocument();
        sst = null;
        stylesTable = null;
        currentCell = null;
        row = null;
    }

    public abstract void getRows(int sheetIndex, Row row);

    public abstract void getLastRowNum(int lastRowNum);
}