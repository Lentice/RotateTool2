package lahome.rotateTool.module;

import com.sun.javafx.scene.control.behavior.TableCellBehavior;
import javafx.scene.control.*;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.util.Callback;
import javafx.util.StringConverter;
import javafx.util.converter.DefaultStringConverter;
import lahome.rotateTool.view.RootController;

/**
 * Created by LenticeTsai on 2017/5/16.
 */
public class DragSelectionCell<S, T> extends TextFieldTableCell<S, T> {

    public DragSelectionCell(StringConverter<T> converter) {
        super(converter);
        initMouseDrag();
    }

    public DragSelectionCell() {
        super();
        initMouseDrag();
    }

    public static <S> Callback<TableColumn<S, String>, TableCell<S, String>> forTableColumn() {
        return forTableColumn(new DefaultStringConverter());
    }

    public static <S, T> Callback<TableColumn<S, T>, TableCell<S, T>> forTableColumn(
            final StringConverter<T> converter) {
        return list -> new DragSelectionCell<>(converter);
    }

    private void initMouseDrag() {
        setOnDragDetected(event -> {
            startFullDrag();
            //log.info("DragDetected");
        });
        setOnMouseDragEntered(event -> {
            performSelection(getTableView(), getTableColumn(), getIndex());
            //log.info("DragEntered");
        });

        setOnMouseReleased(event -> {
            getTableView().getSelectionModel().clearAndSelect(getIndex(), getTableColumn());
        });
    }

    private void performSelection(TableView<S> table, TableColumn<S, T> column, int index) {
        final TablePositionBase anchor = TableCellBehavior.getAnchor(table, table.getFocusModel().getFocusedCell());
        int columnIndex = table.getVisibleLeafIndex(column);

        int minRowIndex = Math.min(anchor.getRow(), index);
        int maxRowIndex = Math.max(anchor.getRow(), index);
        TableColumnBase minColumn = anchor.getColumn() < columnIndex ? anchor.getTableColumn() : column;
        TableColumnBase maxColumn = anchor.getColumn() >= columnIndex ? anchor.getTableColumn() : column;

        table.getSelectionModel().clearSelection();
        final int minColumnIndex = table.getVisibleLeafIndex((TableColumn) minColumn);
        final int maxColumnIndex = table.getVisibleLeafIndex((TableColumn) maxColumn);
        for (int _row = minRowIndex; _row <= maxRowIndex; _row++) {
            for (int _col = minColumnIndex; _col <= maxColumnIndex; _col++) {
                table.getSelectionModel().select(_row, table.getVisibleLeafColumn(_col));
            }
        }

        table.getFocusModel().focus(maxRowIndex, column);
    }
}
