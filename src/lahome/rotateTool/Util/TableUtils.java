package lahome.rotateTool.Util;

import javafx.beans.property.*;
import javafx.beans.value.ObservableValue;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.control.*;
import javafx.scene.input.*;
import lahome.rotateTool.Main;

import java.text.NumberFormat;
import java.text.ParseException;
import java.util.StringTokenizer;

public class TableUtils {
    static UndoManager undoManager;
    private static NumberFormat numberFormatter = NumberFormat.getNumberInstance();

    public static void installMyHandler(TableView<?> table) {

        table.setOnKeyPressed(new TableKeyEventHandler());
        if (undoManager == null) {
            undoManager = UndoManager.getInstance();
        }

//        MenuItem item = new MenuItem("Copy");
//        item.setOnAction(event -> {
//            copySelectionToClipboard(table);
//        });
//
//        ContextMenu menu = new ContextMenu();
//        menu.getItems().add(item);
//        table.setContextMenu(menu);
    }

    /**
     * Copy/Paste keyboard event handler.
     * The handler uses the keyEvent's source for the clipboard data. The source must be of type TableView.
     */
    public static class TableKeyEventHandler implements EventHandler<KeyEvent> {

        KeyCodeCombination copyKeyCodeCompination = new KeyCodeCombination(KeyCode.C, KeyCombination.CONTROL_ANY);
        KeyCodeCombination pasteKeyCodeCompination = new KeyCodeCombination(KeyCode.V, KeyCombination.CONTROL_ANY);
        KeyCodeCombination cutKeyCodeCompination = new KeyCodeCombination(KeyCode.X, KeyCombination.CONTROL_ANY);
        KeyCodeCombination undoKeyCodeCompination = new KeyCodeCombination(KeyCode.Z, KeyCombination.CONTROL_ANY);
        KeyCodeCombination redoKeyCodeCompination = new KeyCodeCombination(KeyCode.Y, KeyCombination.CONTROL_ANY);

        public void handle(final KeyEvent keyEvent) {

            if (copyKeyCodeCompination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    copySelectionToClipboard((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (pasteKeyCodeCompination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    pasteFromClipboard((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (cutKeyCodeCompination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    copySelectionToClipboard((TableView<?>) keyEvent.getSource());
                    deleteSelectCells((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
                }
            } else if (undoKeyCodeCompination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    undoManager.undo();
                }
            } else if (redoKeyCodeCompination.match(keyEvent)) {
                if (keyEvent.getSource() instanceof TableView) {
                    undoManager.redo();
                }
            } else if (keyEvent.getCode() == KeyCode.DELETE) {
                if (keyEvent.getSource() instanceof TableView) {
                    deleteSelectCells((TableView<?>) keyEvent.getSource());
                    keyEvent.consume();
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
                System.out.println("Unsupported observable value: " + observableValue);
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

        // get the cell position to start with
        TablePosition pasteCellPosition = table.getSelectionModel().getSelectedCells().get(0);

        System.out.println("Pasting into cell " + pasteCellPosition);

        String pasteString = Clipboard.getSystemClipboard().getString();

        System.out.println(pasteString);

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

                // System.out.println( rowClipboard + "/" + colClipboard + ": " + cell);

                // get cell
                TableColumn tableColumn = table.getColumns().get(colTable);
                ObservableValue observableValue = tableColumn.getCellObservableValue(rowTable);

                System.out.println(rowTable + "/" + colTable + ": " + observableValue);

                // TODO: handle boolean, etc
                if (observableValue instanceof DoubleProperty) {
                    try {
                        double value = numberFormatter.parse(clipboardCellContent).doubleValue();
                        ((DoubleProperty) observableValue).set(value);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                } else if (observableValue instanceof IntegerProperty) {
                    try {
                        int value = NumberFormat.getInstance().parse(clipboardCellContent).intValue();
                        ((IntegerProperty) observableValue).set(value);
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }
                } else if (observableValue instanceof StringProperty) {
                    ((StringProperty) observableValue).set(clipboardCellContent);
                } else {
                    System.out.println("Unsupported observable value: " + observableValue);
                }

                System.out.println(rowTable + "/" + colTable);
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
                System.out.println("Unsupported observable value: " + observableValue);
            }
        }

        Main.refreshAllTable();
    }
}