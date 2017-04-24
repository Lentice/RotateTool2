package lahome.rotateTool.view;

import com.jfoenix.controls.*;
import com.jfoenix.controls.cells.editors.IntegerTextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.TextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.base.GenericEditableTreeTableCell;
import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.SplitPane;
import javafx.scene.control.Tab;
import javafx.scene.control.TreeItem;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.stage.FileChooser;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.ExcelParser;
import lahome.rotateTool.module.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.util.prefs.Preferences;

public class RootController {
    private static final Logger log = LogManager.getLogger(RootController.class.getName());

    @FXML
    private JFXTabPane mainTab;

    @FXML
    private Tab fileTab;

    @FXML
    private Tab editTab;

    @FXML
    private SplitPane editTabSlitPane;

    @FXML
    private JFXTextField pathAgingReport;

    @FXML
    private JFXTextField pathStockFile;

    @FXML
    private JFXTextField pathPurchaseFile;

    @FXML
    private ImageView filterIcon;

    @FXML
    private JFXTextField filterField;

    @FXML
    private JFXTreeTableView<RotateItem> rotateTreeTableView;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateKitColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotatePartColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotatePmQtyColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotateMyQtyColumn;

    @FXML
    private JFXTreeTableView<StockItem> stockTreeTableView;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockPoColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockQtyColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockDcColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockMyQtyColumn;


    private Main main;
    private RotateCollection collection;
    private File latestDir = new File(System.getProperty("user.dir"));

    public RootController() {
    }

    public void saveSetting() {
        String path;
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        prefs.put("LatestDir", latestDir.getPath());

        path = pathAgingReport.getText();
        prefs.put("AgingPath", path != null ? path : "");
        path = pathStockFile.getText();
        prefs.put("StockPath", path != null ? path : "");
        path = pathPurchaseFile.getText();
        prefs.put("PurchasePath", path != null ? path : "");

        prefs.put("EditTabDiv", String.valueOf(editTabSlitPane.getDividerPositions()[0]));

        prefs.put("RotateKitWidth", String.valueOf(rotateKitColumn.getWidth()));
        prefs.put("RotatePartWidth", String.valueOf(rotatePartColumn.getWidth()));
        prefs.put("RotatePmQtyWidth", String.valueOf(rotatePmQtyColumn.getWidth()));
        prefs.put("RotateMyQtyWidth", String.valueOf(rotateMyQtyColumn.getWidth()));

        prefs.put("StockPoWidth", String.valueOf(stockPoColumn.getWidth()));
        prefs.put("StockDcWidth", String.valueOf(stockDcColumn.getWidth()));
        prefs.put("StockQtyWidth", String.valueOf(stockQtyColumn.getWidth()));
        prefs.put("StockMyQtyWidth", String.valueOf(stockMyQtyColumn.getWidth()));
    }

    public void loadSetting() {
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        latestDir = new File(prefs.get("LatestDir", System.getProperty("user.dir")));
        if (!latestDir.exists())
            latestDir = new File(System.getProperty("user.dir"));

        pathAgingReport.setText(prefs.get("AgingPath", ""));
        pathStockFile.setText(prefs.get("StockPath", ""));
        pathPurchaseFile.setText(prefs.get("PurchasePath", ""));

        editTabSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabDiv", "0.41")));

        rotateKitColumn.setPrefWidth(Double.valueOf(prefs.get("RotateKitWidth", "140")));
        rotatePartColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePartWidth", "140")));
        rotatePmQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePmQtyWidth", "50")));
        rotateMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotateMyQtyWidth", "50")));

        stockPoColumn.setPrefWidth(Double.valueOf(prefs.get("StockPoWidth", "120")));
        stockDcColumn.setPrefWidth(Double.valueOf(prefs.get("StockDcWidth", "50")));
        stockQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockQtyWidth", "63")));
        stockMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockMyQtyWidth", "50")));
    }

    public void setMainApp(Main main) {
        this.main = main;
        this.collection = main.getCollection();
    }


    @FXML
    void handleBrowseAgingReport() {
        FileChooser fileChooser = new FileChooser();

        fileChooser.setTitle("Select Aging Report");
        fileChooser.setInitialDirectory(latestDir);
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel file", "*.xlsx;*.xls"));

        File file = fileChooser.showOpenDialog(main.getPrimaryStage());
        if (file != null) {
            pathAgingReport.setText(file.getPath());
            latestDir = new File(file.getParent());
        }
    }

    @FXML
    void handleBrowseStockFile() {
        FileChooser fileChooser = new FileChooser();

        fileChooser.setTitle("Select Stock Report");
        fileChooser.setInitialDirectory(latestDir);
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel file", "*.xlsx;*.xls"));

        File file = fileChooser.showOpenDialog(main.getPrimaryStage());
        if (file != null) {
            pathStockFile.setText(file.getPath());
            latestDir = new File(file.getParent());
        }
    }

    @FXML
    void handleBrowsePurchaseFile() {
        FileChooser fileChooser = new FileChooser();

        fileChooser.setTitle("Select Purchase Report");
        fileChooser.setInitialDirectory(latestDir);
        fileChooser.getExtensionFilters().addAll(
                new FileChooser.ExtensionFilter("Excel file", "*.xlsx;*.xls"));

        File file = fileChooser.showOpenDialog(main.getPrimaryStage());
        if (file != null) {
            pathPurchaseFile.setText(file.getPath());
            latestDir = new File(file.getParent());
        }
    }

    @FXML
    void handleStartEdit() {
        File file;
        ExcelParser excelParser = new ExcelParser(collection);

        file = new File(pathAgingReport.getText());
        excelParser.loadRotateExcel(file, 4, "F", "G", "R");
        log.info("Process rotate table: Done");

        file = new File(pathStockFile.getText());

        excelParser.loadStockExcel(file, 2, "G", "F", "I",
                "J", "P", "L");
        log.info("Process stock table: Done");

        final TreeItem<RotateItem> rotateRoot = new RecursiveTreeItem<>(
                collection.getRotateObsList(), RecursiveTreeObject::getChildren);
        rotateTreeTableView.setRoot(rotateRoot);
        rotateTreeTableView.getSelectionModel().select(0);
        rotateTreeTableView.requestFocus();

//        rotateTreeTableView.group(rotateKitColumn);
//        for(TreeItem<RotateItem> child:rotateTreeTableView.getRoot().getChildren()){
//            child.setExpanded(true);
//        }

        mainTab.getSelectionModel().select(editTab);
    }

    @FXML
    private void initialize() {

        filterIcon.setImage(new Image("file:resources/images/active-search-2-48.png"));

        initialRotateTable();
        initialStockTable();

    }

    private void initialRotateTable() {
        rotateKitColumn.setStyle("-fx-alignment: top-left; ");
        rotatePartColumn.setStyle("-fx-alignment: top-left; ");

        rotateKitColumn.setCellValueFactory(cellData -> {
            if (rotateKitColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().kitNameProperty();
            } else {
                return rotateKitColumn.getComputedValue(cellData);
            }
        });

        rotatePartColumn.setCellValueFactory(cellData -> {
            if (rotatePartColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().partNumberProperty();
            } else {
                return rotatePartColumn.getComputedValue(cellData);
            }
        });

        rotatePmQtyColumn.setCellValueFactory(cellData -> {
            if (rotatePmQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().pmQtyProperty();
            } else {
                return rotatePmQtyColumn.getComputedValue(cellData);
            }
        });

        rotateMyQtyColumn.setCellValueFactory(cellData -> {
            if (rotateMyQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().myQtyProperty();
            } else {
                return rotateMyQtyColumn.getComputedValue(cellData);
            }
        });

        rotateKitColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        rotateKitColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().kitNameProperty().set(cellData.getNewValue());
            rotateTreeTableView.requestFocus();
        });

        rotatePartColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        rotatePartColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().partNumberProperty().set(cellData.getNewValue());
            rotateTreeTableView.requestFocus();
        });

        filterField.textProperty().addListener((o, oldVal, newVal) -> {
            rotateTreeTableView.setPredicate(itemProp -> {
                final RotateItem item = itemProp.getValue();
                return item.kitNameProperty().get().contains(newVal) || item.partNumberProperty().get().contains(newVal);
            });
        });

        rotateTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedRotateItem(oldValue, newValue)
        );


        rotateTreeTableView.setRowFactory(tv -> new JFXTreeTableRow<RotateItem>() {
            @Override
            public void updateItem(RotateItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null) {
                    setStyle("");
                } else if (item.isDuplicate()) {
                    setStyle("-fx-background-color: #FFEE58;");
                } else if (item.getPmQty() == item.getMyQty()) {
                    setStyle("-fx-background-color: #4CAF50;");
                } else if (item.getPmQty() < item.getMyQty()) {
                    setStyle("-fx-background-color: #ef5350;");
                } else {
                    setStyle("");
                }
            }
        });

    }

    private void initialStockTable() {
        stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
        stockPoColumn.setStyle("-fx-alignment: top-left; ");

        stockPoColumn.setCellValueFactory(cellData -> {
            if (stockPoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return stockPoColumn.getComputedValue(cellData);
            }
        });

        stockQtyColumn.setCellValueFactory(cellData -> {
            if (stockQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().stockQtyProperty();
            } else {
                return stockQtyColumn.getComputedValue(cellData);
            }
        });

        stockDcColumn.setCellValueFactory(cellData -> {
            if (stockDcColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().dcProperty();
            } else {
                return stockDcColumn.getComputedValue(cellData);
            }
        });

        stockMyQtyColumn.setCellValueFactory(cellData -> {
            if (stockMyQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().myQtyProperty();
            } else {
                return stockMyQtyColumn.getComputedValue(cellData);
            }
        });

        stockMyQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        stockMyQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().myQtyProperty().set(cellData.getNewValue().intValue());
            stockTreeTableView.requestFocus();
        });

    }


    private void selectedRotateItem(TreeItem<RotateItem> oldValue, TreeItem<RotateItem> newValue) {

        if (newValue == null || !newValue.isLeaf()) {
            stockTreeTableView.setRoot(null);
            return;
        }

        RotateItem rotateItem = newValue.getValue();
        ObservableList<StockItem> obsList = rotateItem.getStockItemObsList();
        TreeItem<StockItem> root = new RecursiveTreeItem<>(
                rotateItem.getStockItemObsList(), RecursiveTreeObject::getChildren);

        if (rotateItem.isDuplicate()) {
            root = null;
            stockTreeTableView.setPlaceholder(new Label("重複的項目!! 我無法幫妳了 orz...."));
        } else if (obsList.size() == 0) {
            stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
            root = null;
        }

        stockTreeTableView.setRoot(root);
    }
}
