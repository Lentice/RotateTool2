package lahome.rotateTool.module;

import com.sun.javafx.scene.control.behavior.TableCellBehavior;
import javafx.application.Platform;
import javafx.beans.value.ChangeListener;
import javafx.event.Event;
import javafx.event.EventHandler;
import javafx.scene.control.*;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyEvent;
import javafx.scene.input.MouseButton;
import javafx.util.Callback;
import javafx.util.StringConverter;
import javafx.util.converter.DefaultStringConverter;

public class DragSelectionCell<S, T> extends TextFieldTableCell<S, T> {

//    ChangeListener<? super Boolean> changeListener;

    @SuppressWarnings("WeakerAccess")
    public DragSelectionCell(StringConverter<T> converter) {
        super(converter);
        initMouseDrag();
    }

    @SuppressWarnings("WeakerAccess")
    public DragSelectionCell() {
        super();
        initMouseDrag();

    }

//    @Override
//    public void startEdit() {
//        super.startEdit();
//
//        changeListener = (observable, oldSelection, newSelection) ->
//        {
//            if (!newSelection) {
//                commitEdit(getItem());
//            }
//        };
//        focusedProperty().addListener(changeListener);
//    }

//    @Override
//    public void commitEdit(T newValue) {
//        super.commitEdit(newValue);
//
//        if (changeListener != null) {
//            focusedProperty().removeListener(changeListener);
//            changeListener = null;
//        }
//    }

    public static <S> Callback<TableColumn<S, String>, TableCell<S, String>> forTableColumn() {
        return forTableColumn(new DefaultStringConverter());
    }

    public static <S, T> Callback<TableColumn<S, T>, TableCell<S, T>> forTableColumn(
            final StringConverter<T> converter) {
        return list -> new DragSelectionCell<>(converter);
    }

    private void initMouseDrag() {
        setOnDragDetected(event -> {
            if (event.getButton() == MouseButton.PRIMARY) {
                startFullDrag();
            }
            //log.info("DragDetected");
        });
        setOnMouseDragEntered(event -> {
            if (event.getButton() == MouseButton.PRIMARY) {
                performSelection(getTableView(), getTableColumn(), getIndex());
            }
            //log.info("DragEntered");
        });

        setOnMouseReleased(event -> {
            if (event.getButton() == MouseButton.PRIMARY) {
                getTableView().getSelectionModel().clearAndSelect(getIndex(), getTableColumn());
            }
        });
    }

    private void performSelection(TableView<S> table, TableColumn<S, T> column, int index) {
        final TablePositionBase anchor = TableCellBehavior.getAnchor(table, table.getFocusModel().getFocusedCell());
        int columnIndex = table.getVisibleLeafIndex(column);

        int minRowIndex = Math.min(anchor.getRow(), index);
        int maxRowIndex = Math.max(anchor.getRow(), index);
        TableColumnBase minColumn = anchor.getColumn() < columnIndex ? anchor.getTableColumn() : column;
        TableColumnBase maxColumn = anchor.getColumn() >= columnIndex ? anchor.getTableColumn() : column;

        @SuppressWarnings("unchecked") final int minColumnIndex = table.getVisibleLeafIndex((TableColumn<S, ?>) minColumn);
        @SuppressWarnings("unchecked") final int maxColumnIndex = table.getVisibleLeafIndex((TableColumn<S, ?>) maxColumn);

        table.getSelectionModel().clearSelection();
        for (int _row = minRowIndex; _row <= maxRowIndex; _row++) {
            for (int _col = minColumnIndex; _col <= maxColumnIndex; _col++) {
                table.getSelectionModel().select(_row, table.getVisibleLeafColumn(_col));
            }
        }

        table.getFocusModel().focus(index, column);
    }
}
