package lahome.rotateTool.Util;

import javafx.beans.property.DoubleProperty;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.Property;
import javafx.beans.property.StringProperty;
import javafx.scene.control.TablePosition;

import java.util.Stack;

/**
 * Created by Administrator on 2017/5/15.
 */
public class UndoManager {

    private static UndoManager instance = null;
    private final Stack<Command> undoStack = new Stack<>();
    private final Stack<Command> redoStack = new Stack<>();
    private static int commandTag = 0;

    public void newInput() {
        commandTag++;
        redoStack.clear();
    }

    public void inputHandle(Property property, Object newValue, TablePosition<?, ?> tablePosition) {
        Command cmd = new Command(commandTag, property, newValue, tablePosition);
        undoStack.push(cmd);
        cmd.execute();
    }

    public void newInput(Property property, Object newValue) {
        newInput();
        inputHandle(property, newValue, null);
    }

    public void newInput(Property property, Object newValue, TablePosition<?, ?> position) {
        newInput();
        inputHandle(property, newValue, position);
    }

    public void appendInput(Property property, Object newValue) {
        inputHandle(property, newValue, null);
    }

    public void appendInput(Property property, Object newValue, TablePosition<?, ?> position) {
        inputHandle(property, newValue, position);
    }

    public void undo() {
        int currentTag = -1;
        while (!undoStack.isEmpty()) {
            Command cmd = undoStack.pop();
            if (currentTag == -1)
                currentTag = cmd.getTag();
            else if (currentTag != cmd.getTag()) {
                undoStack.push(cmd);
                return;
            }

            cmd.undo();
            redoStack.push(cmd);
        }

        if (currentTag != -1)
            commandTag = currentTag;
    }

    public void redo() {
        int currentTag = -1;
        while (!redoStack.isEmpty()) {
            Command cmd = redoStack.pop();
            if (currentTag == -1)
                currentTag = cmd.getTag();
            else if (currentTag != cmd.getTag()) {
                redoStack.push(cmd);
                return;
            }

            cmd.execute();
            undoStack.push(cmd);
        }

        if (currentTag != -1) {
            commandTag = currentTag + 1;
        }
    }

    public static UndoManager getInstance() {
        if (UndoManager.instance == null) {
            synchronized (UndoManager.class) {
                if (UndoManager.instance == null) {
                    UndoManager.instance = new UndoManager();
                }
            }
        }
        return UndoManager.instance;
    }


    public static class Command {
        int tag;
        Object observableValue;
        Object oldValue;
        Object newValue;
        private TablePosition<?, ?> tablePostion;

        Command(int tag, Object observableValue, Object newValue) {
            this.tag = tag;
            this.observableValue = observableValue;
            this.newValue = newValue;

            if (observableValue instanceof DoubleProperty) {
                oldValue = ((DoubleProperty) observableValue).get();
            } else if (observableValue instanceof IntegerProperty) {
                oldValue = ((IntegerProperty) observableValue).get();
            } else if (observableValue instanceof StringProperty) {
                oldValue = ((StringProperty) observableValue).get();
            } else {
                System.out.println("Unsupported observable value: " + observableValue);
            }
        }

        Command(int tag, Object observableValue, Object newValue, TablePosition<?, ?> tablePosition) {
            this(tag, observableValue, newValue);
            this.tablePostion = tablePosition;
        }

        void undo() {
            if (tablePostion != null) {
                tablePostion.getTableView().requestFocus();
                tablePostion.getTableView().scrollTo(tablePostion.getRow());
                tablePostion.getTableView().scrollToColumnIndex(tablePostion.getColumn());
                tablePostion.getTableView().getSelectionModel().clearAndSelect(
                        tablePostion.getRow()
                );
                try {
                    Thread.sleep(50);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }

            if (observableValue instanceof DoubleProperty) {
                ((DoubleProperty) observableValue).set((Double) oldValue);
            } else if (observableValue instanceof IntegerProperty) {
                ((IntegerProperty) observableValue).set((Integer) oldValue);

            } else if (observableValue instanceof StringProperty) {
                ((StringProperty) observableValue).set((String) oldValue);
            } else {
                System.out.println("Unsupported observable value: " + observableValue);
            }
        }

        void execute() {
            if (observableValue instanceof DoubleProperty) {
                ((DoubleProperty) observableValue).set((Double) newValue);
            } else if (observableValue instanceof IntegerProperty) {
                ((IntegerProperty) observableValue).set((Integer) newValue);
            } else if (observableValue instanceof StringProperty) {
                ((StringProperty) observableValue).set((String) newValue);
            } else {
                System.out.println("Unsupported observable value: " + observableValue);
            }
        }

        int getTag() {
            return tag;
        }
    }
}
