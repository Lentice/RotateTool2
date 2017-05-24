package lahome.rotateTool.Util;

import javafx.beans.property.*;
import javafx.beans.value.ObservableValue;
import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.*;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.Excel.ExcelReader;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.text.NumberFormat;
import java.text.ParseException;
import java.util.StringTokenizer;

public class TableUtils {
    private static final Logger log = LogManager.getLogger(TableUtils.class.getName());

    static UndoManager undoManager;
    private static NumberFormat numberFormatter = NumberFormat.getNumberInstance();

    private String lastKey = null;

    public static void installMyHandler(TableView<?> table) {

        if (undoManager == null) {
            undoManager = UndoManager.getInstance();
        }

        table.setOnKeyPressed(new TableKeyEventHandler());
        setContextMenu(table);
    }

    private static void setContextMenu(TableView<?> table) {
        ContextMenu menu = new ContextMenu();
        table.setContextMenu(menu);

        MenuItem copy = new MenuItem("Copy", new ImageView(
                new Image(Main.class.getResourceAsStream("/images/Copy_24px.png"))));
        copy.setOnAction(event -> {
            copySelectionToClipboard(table);
        });
        menu.getItems().add(copy);

        MenuItem cut = new MenuItem("Cut", new ImageView(
                new Image(Main.class.getResourceAsStream("/images/Cut_24px.png"))));
        cut.setOnAction(event -> {
            copySelectionToClipboard(table);
            deleteSelectCells(table);
        });
        menu.getItems().add(cut);

        MenuItem paste = new MenuItem("Paste", new ImageView(
                new Image(Main.class.getResourceAsStream("/images/Paste_24px.png"))));
        paste.setOnAction(event -> {
            pasteFromClipboard(table);
        });
        menu.getItems().add(paste);
    }

    /**
     * Copy/Paste keyboard event handler.
     * The handler uses the keyEvent's source for the clipboard data. The source must be of type TableView.
     */
    public static class TableKeyEventHandler implements EventHandler<KeyEvent> {

        KeyCodeCombination copyKeyCodeCombination = new KeyCodeCombination(KeyCode.C, KeyCombination.CONTROL_ANY);
        KeyCodeCombination pasteKeyCodeCombination = new KeyCodeCombination(KeyCode.V, KeyCombination.CONTROL_ANY);
        KeyCodeCombination cutKeyCodeCombination = new KeyCodeCombination(KeyCode.X, KeyCombination.CONTROL_ANY);
        KeyCodeCombination undoKeyCodeCombination = new KeyCodeCombination(KeyCode.Z, KeyCombination.CONTROL_ANY);
        KeyCodeCombination redoKeyCodeCombination = new KeyCodeCombination(KeyCode.Y, KeyCombination.CONTROL_ANY);

        public void handle(final KeyEvent keyEvent) {

            if (copyKeyCodeCombination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    copySelectionToClipboard((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (pasteKeyCodeCombination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    pasteFromClipboard((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (cutKeyCodeCombination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    copySelectionToClipboard((TableView<?>) keyEvent.getSource());
                    deleteSelectCells((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (undoKeyCodeCombination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    undoManager.undo();
                }
            } else if (redoKeyCodeCombination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    undoManager.redo();
                }
            } else if (keyEvent.getCode() == KeyCode.DELETE) {
                if (keyEvent.getSource() instanceof TableView) {
                    deleteSelectCells((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else {
                TablePosition tp;
                if (keyEvent.getSource() instanceof TableView) {
                    if (!keyEvent.isControlDown() &&
                            (keyEvent.getCode().isLetterKey() || keyEvent.getCode().isDigitKey())) {

                        TableView<?> table = (TableView<?>) keyEvent.getSource();
                        tp = table.getFocusModel().getFocusedCell();
                        //noinspection unchecked
                        table.edit(tp.getRow(), tp.getTableColumn());
                    }
                }
            }
        }
    }

    /**
     * Get table selection and copy it to the clipboard.
     *
     * @param table
     */
    public static void copySelectionToClipboard(TableView<?> table) {

        StringBuilder clipboardString = new StringBuilder();

        ObservableList<TablePosition> positionList = table.getSelectionModel().getSelectedCells();

        int prevRow = -1;

        for (TablePosition position : positionList) {

            int row = position.getRow();
            int col = position.getColumn();

            // determine whether we advance in a row (tab) or a column
            // (newline).
            if (prevRow == row) {
                clipboardString.append('\t');
            } else if (prevRow != -1) {
                clipboardString.append('\n');
            }

            // create string from cell
            String text = "";

            Object observableValue = (Object) table.getColumns().get(col).getCellObservableValue(row);

            // null-check: provide empty string for nulls
            if (observableValue == null) {
                text = "";
            } else if (observableValue instanceof DoubleProperty) { // TODO: handle boolean etc
                text = numberFormatter.format(((DoubleProperty) observableValue).get());
            } else if (observableValue instanceof IntegerProperty) {
                text = numberFormatter.format(((IntegerProperty) observableValue).get());
            } else if (observableValue instanceof StringProperty) {
                text = ((StringProperty) observableValue).get();
            } else if (observableValue instanceof ReadOnlyDoubleProperty) {
                text = numberFormatter.format(((ReadOnlyDoubleProperty) observableValue).get());
            } else if (observableValue instanceof ReadOnlyIntegerProperty) {
                text = numberFormatter.format(((ReadOnlyIntegerProperty) observableValue).get());
            } else if (observableValue instanceof ReadOnlyStringProperty) {
                text = ((ReadOnlyStringProperty) observableValue).get();
            } else {
                log.error("Unsupported observable value: " + observableValue);
            }

            // add new item to clipboard
            clipboardString.append(text);

            // remember previous
            prevRow = row;
        }

        // create clipboard content
        final ClipboardContent clipboardContent = new ClipboardContent();
        clipboardContent.putString(clipboardString.toString());

        // set clipboard content
        Clipboard.getSystemClipboard().setContent(clipboardContent);


    }

    public static void pasteFromClipboard(TableView<?> table) {

        // abort if there's not cell selected to start with
        if (table.getSelectionModel().getSelectedCells().size() == 0) {
            return;
        }

        undoManager.newInput();

        // get the cell position to start with
        TablePosition pasteCellPosition = table.getSelectionModel().getSelectedCells().get(0);

        String pasteString = Clipboard.getSystemClipboard().getString();

        int rowClipboard = -1;

        StringTokenizer rowTokenizer = new StringTokenizer(pasteString, "\n");
        while (rowTokenizer.hasMoreTokens()) {

            rowClipboard++;

            String rowString = rowTokenizer.nextToken();

            StringTokenizer columnTokenizer = new StringTokenizer(rowString, "\t");

            int colClipboard = -1;

            while (columnTokenizer.hasMoreTokens()) {

                colClipboard++;

                // get next cell data from clipboard
                String clipboardCellContent = columnTokenizer.nextToken();

                // calculate the position in the table cell
                int rowTable = pasteCellPosition.getRow() + rowClipboard;
                int colTable = pasteCellPosition.getColumn() + colClipboard;

                // skip if we reached the end of the table
                if (rowTable >= table.getItems().size()) {
                    continue;
                }
                if (colTable >= table.getColumns().size()) {
                    continue;
                }

                // get cell
                TableColumn tableColumn = table.getColumns().get(colTable);
                ObservableValue observableValue = tableColumn.getCellObservableValue(rowTable);

                // TODO: handle boolean, etc
                try {
                    if (observableValue instanceof DoubleProperty) {

                        double value = numberFormatter.parse(clipboardCellContent).doubleValue();
                        undoManager.appendInput((DoubleProperty) observableValue, value);

                    } else if (observableValue instanceof IntegerProperty) {
                        int value = NumberFormat.getInstance().parse(clipboardCellContent).intValue();
                        undoManager.appendInput((IntegerProperty) observableValue, value);
                    } else if (observableValue instanceof StringProperty) {
                        undoManager.appendInput((StringProperty) observableValue, clipboardCellContent);
                    } else {
                        log.error("Unsupported observable value: " + observableValue);
                    }

                } catch (ParseException e) {
                    log.error("failed", e);
                }
            }
        }
    }


    public static void deleteSelectCells(TableView<?> table) {

        // abort if there's not cell selected to start with
        if (table.getSelectionModel().getSelectedCells().size() == 0) {
            return;
        }

        undoManager.newInput();

        ObservableList<TablePosition> positionList = table.getSelectionModel().getSelectedCells();
        for (TablePosition position : positionList) {

            int row = position.getRow();
            int col = position.getColumn();

            ObservableValue observableValue = table.getColumns().get(col).getCellObservableValue(row);
            if (observableValue instanceof DoubleProperty) {
                undoManager.appendInput(((DoubleProperty) observableValue), 0);
            } else if (observableValue instanceof IntegerProperty) {
                undoManager.appendInput(((IntegerProperty) observableValue), 0);
            } else if (observableValue instanceof StringProperty) {
                undoManager.appendInput(((StringProperty) observableValue), "");
            } else {
                log.error("Unsupported observable value: " + observableValue);
            }
        }

        Main.refreshAllTable();
    }
}