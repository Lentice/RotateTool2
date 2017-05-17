package lahome.rotateTool.Util.Excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
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

@SuppressWarnings("WeakerAccess")
public abstract class StreamingReader extends DefaultHandler {

    // 共享字符串表
    private SharedStringsTable sst;
    private StylesTable stylesTable;

    private final DataFormatter dataFormatter = new DataFormatter();

    private boolean useDate1904 = false;

    private StringBuilder lastContents = new StringBuilder(1024);

    private int sheetIndex = -1;
    private int curRow = 0;
    private int curCol = 0;
    private String cellStyleString;
    private StreamingCell currentCell;
    private StreamingRow row = new StreamingRow(0);

    public void processAllSheets(String filename) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        stylesTable = r.getStylesTable();
        sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser();
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
        IOUtils.closeQuietly(pkg);
    }

    public void processOneSheet(String filename, int sheetIndex) throws Exception {
        OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
        XSSFReader r = new XSSFReader(pkg);
        stylesTable = r.getStylesTable();
        sst = r.getSharedStringsTable();

        XMLReader parser = fetchSheetParser();

        // To look up the Sheet Name / Sheet Order / rID,
        //  you need to processOneSheet the core Workbook stream.
        // Normally it's of the form rId# or rSheet#
        InputStream sheet = r.getSheet("rId" + sheetIndex);
        this.sheetIndex++;
        InputSource sheetSource = new InputSource(sheet);
        parser.parse(sheetSource);

        sheet.close();
        pkg.revert();
        IOUtils.closeQuietly(pkg);
    }


    public XMLReader fetchSheetParser() throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
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

    private String formattedContents(String content) {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                int idx = Integer.parseInt(content);
                return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(content).toString();
            case "str":         //forumla type
                return '"' + content + '"';
            case "e":           //error type
                return "ERROR:  " + content;
            case "n":           //numeric type
//                if(currentCell.getNumericFormat() != null && content.length() > 0) {
//                    return dataFormatter.formatRawCellContents(
//                            Double.parseDouble(content),
//                            currentCell.getNumericFormatIndex(),
//                            currentCell.getNumericFormat());
//                } else {
                return content;
//                }
            default:
                return content;
        }
    }

    private String unformattedContents(String content) {
        switch (currentCell.getType()) {
            case "s":           //string stored in shared table
                int idx = Integer.parseInt(content);
                return new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            case "inlineStr":   //inline string (not in sst)
                return new XSSFRichTextString(content).toString();
            default:
                return content;
        }
    }

    private void getCoordinate(String coord) {
        int col = 0;
        int row = 0;

        for (int i = 0; i < coord.length(); i++) {
            char ch = coord.charAt(i);
            if (!isNumber(ch)) {
                col = col * 26 + ch - 65 + 1;
            } else if (curCol == 0) {
                row = row * 10 + ch - 48;
            }
        }

        if (curCol == 0)
            curRow = row - 1;

        curCol = col - 1;
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
            currentCell = row.createCell(curCol);
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
                row.setRowNum(curRow);
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
        lastContents.setLength(0);
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {

        if ("v".equals(name) || "t".equals(name)) {
            String content = lastContents.toString();
            currentCell.setRawContents(unformattedContents(content));
            //currentCell.setContents(formattedContents(content));

        } else if ("row".equals(name)) {
            getRows(sheetIndex, row);
            //row = null;
            curCol = 0;
            curRow++;
            row.clear();
            row.setRowNum(curRow);
        } else if ("c".equals(name)) {
            //if (row == null) {
            //    row = new StreamingRow(curRow);
            //}
            //row.addCell(curCol, currentCell);
            curCol++;
        } else if ("f".equals(name)) {
            currentCell.setFormula(lastContents.toString());
        }
    }

    public void characters(char[] ch, int start, int length)
            throws SAXException {
        lastContents.append(ch, start, length);
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