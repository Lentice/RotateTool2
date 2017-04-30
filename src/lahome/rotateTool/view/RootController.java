package lahome.rotateTool.view;

import com.jfoenix.controls.*;
import com.jfoenix.controls.cells.editors.IntegerTextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.TextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.base.GenericEditableTreeTableCell;
import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.stage.FileChooser;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.DateUtil;
import lahome.rotateTool.Util.ExcelParser;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.time.LocalDate;
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
    private SplitPane editTabHSlitPane;

    @FXML
    private SplitPane editTabVLeftSlitPane;

    @FXML
    private SplitPane editTabVRightSlitPane;

    @FXML
    private DatePicker grDateStart;

    @FXML
    private DatePicker grDateEnd;

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
    private JFXTreeTableColumn<RotateItem, String> rotateRatioColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotateApplySetColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateRemarkColumn;

    @FXML
    private JFXTreeTableView<StockItem> stockTreeTableView;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockPoColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockLotColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockDcColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockQtyColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockMyQtyColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockRemarkColumn;

    @FXML
    private Label stockStQtyTotal;

    @FXML
    private Label stockGrQtyTotal;

    @FXML
    private Label stockApQtyTotal;

    @FXML
    private Label stockCurrPoStQtyTotal;

    @FXML
    private Label stockCurrPoApQtyTotal;

    @FXML
    private JFXTreeTableView<PurchaseItem> purchaseTreeTableView;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchaseGrDateColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> purchaseGrQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> purchaseMyQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> purchaseApplySetColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchaseRemarkColumn;

    @FXML
    private Label purchaseGrQtyTotal;

    @FXML
    private Label purchaseApQtyTotal;

    @FXML
    private Label purchaseApSetTotal;

    @FXML
    private JFXTreeTableView<PurchaseItem> noneStPurchaseTreeTableView;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchasePoColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchaseGrDateColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> noneStPurchaseGrQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> noneStPurchaseMyQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> noneStPurchaseApplySetColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchaseRemarkColumn;

    @FXML
    private Label noneStPurchaseGrQtyTotal;

    @FXML
    private Label noneStPurchaseApQtyTotal;

    @FXML
    private Label noneStPurchaseApSetTotal;


    private Main main;
    private RotateCollection collection;
    private File latestDir = new File(System.getProperty("user.dir"));
    private String rotateSelectedKit = "";
    private String stockSelectedPo = "";
    private String purchaseSelectedPo = "";
    private String noneStPurchaseSelectedPo = "";

    public RootController() {
    }

    public void saveSetting() {
        String path;
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        String dateStart = DateUtil.formatFromLocalDate(grDateStart.getValue());
        prefs.put("GrDateStart", dateStart == null ? "" : dateStart);

        String dateEnd = DateUtil.formatFromLocalDate(grDateEnd.getValue());
        prefs.put("GrDateEnd", dateEnd == null ? "" : dateEnd);

        prefs.put("LatestDir", latestDir.getPath());

        path = pathAgingReport.getText();
        prefs.put("AgingPath", path != null ? path : "");
        path = pathStockFile.getText();
        prefs.put("StockPath", path != null ? path : "");
        path = pathPurchaseFile.getText();
        prefs.put("PurchasePath", path != null ? path : "");

        prefs.put("EditTabHDiv", String.valueOf(editTabHSlitPane.getDividerPositions()[0]));
        prefs.put("EditTabVLeftDiv", String.valueOf(editTabVLeftSlitPane.getDividerPositions()[0]));
        prefs.put("EditTabVRightDiv", String.valueOf(editTabVRightSlitPane.getDividerPositions()[0]));

        prefs.put("RotateKitWidth", String.valueOf(rotateKitColumn.getWidth()));
        prefs.put("RotatePartWidth", String.valueOf(rotatePartColumn.getWidth()));
        prefs.put("RotatePmQtyWidth", String.valueOf(rotatePmQtyColumn.getWidth()));
        prefs.put("RotateMyQtyWidth", String.valueOf(rotateMyQtyColumn.getWidth()));
        prefs.put("RotateRatioWidth", String.valueOf(rotateRatioColumn.getWidth()));
        prefs.put("RotateApplySetWidth", String.valueOf(rotateApplySetColumn.getWidth()));
        prefs.put("RotateRemarkWidth", String.valueOf(rotateRemarkColumn.getWidth()));

        prefs.put("StockPoWidth", String.valueOf(stockPoColumn.getWidth()));
        prefs.put("StockLotWidth", String.valueOf(stockLotColumn.getWidth()));
        prefs.put("StockDcWidth", String.valueOf(stockDcColumn.getWidth()));
        prefs.put("StockQtyWidth", String.valueOf(stockQtyColumn.getWidth()));
        prefs.put("StockMyQtyWidth", String.valueOf(stockMyQtyColumn.getWidth()));
        prefs.put("StockRemarkWidth", String.valueOf(stockRemarkColumn.getWidth()));

        prefs.put("PurchaseGrDateWidth", String.valueOf(purchaseGrDateColumn.getWidth()));
        prefs.put("PurchaseGrQtyWidth", String.valueOf(purchaseGrQtyColumn.getWidth()));
        prefs.put("PurchaseMyQtyWidth", String.valueOf(purchaseMyQtyColumn.getWidth()));
        prefs.put("PurchaseApplySetWidth", String.valueOf(purchaseApplySetColumn.getWidth()));
        prefs.put("PurchaseRemarkWidth", String.valueOf(purchaseRemarkColumn.getWidth()));

        prefs.put("NoneStPurchasePoWidth", String.valueOf(noneStPurchasePoColumn.getWidth()));
        prefs.put("NoneStPurchaseGrDateWidth", String.valueOf(noneStPurchaseGrDateColumn.getWidth()));
        prefs.put("NoneStPurchaseGrQtyWidth", String.valueOf(noneStPurchaseGrQtyColumn.getWidth()));
        prefs.put("NoneStPurchaseMyQtyWidth", String.valueOf(noneStPurchaseMyQtyColumn.getWidth()));
        prefs.put("NoneStPurchaseApplySetWidth", String.valueOf(noneStPurchaseApplySetColumn.getWidth()));
        prefs.put("NoneStPurchaseRemarkWidth", String.valueOf(noneStPurchaseRemarkColumn.getWidth()));
    }

    public void loadSetting() {
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        LocalDate dateStart = DateUtil.parseToLocalDate(prefs.get("GrDateStart", ""));
        if (dateStart != null) {
            grDateStart.setValue(dateStart);
        } else {
            grDateStart.setValue(LocalDate.now());
        }

        LocalDate dateEnd = DateUtil.parseToLocalDate(prefs.get("GrDateEnd", ""));
        if (dateEnd != null) {
            grDateEnd.setValue(dateEnd);
        } else {
            grDateEnd.setValue(LocalDate.now());
        }

        latestDir = new File(prefs.get("LatestDir", System.getProperty("user.dir")));
        if (!latestDir.exists())
            latestDir = new File(System.getProperty("user.dir"));

        pathAgingReport.setText(prefs.get("AgingPath", ""));
        pathStockFile.setText(prefs.get("StockPath", ""));
        pathPurchaseFile.setText(prefs.get("PurchasePath", ""));

        editTabHSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabHDiv", "0.41")));
        editTabVLeftSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabVLeftDiv", "0.75")));
        editTabVRightSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabVRightDiv", "0.75")));

        rotateKitColumn.setPrefWidth(Double.valueOf(prefs.get("RotateKitWidth", "140")));
        rotatePartColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePartWidth", "140")));
        rotatePmQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePmQtyWidth", "50")));
        rotateMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotateMyQtyWidth", "50")));
        rotateRatioColumn.setPrefWidth(Double.valueOf(prefs.get("RotateRatioWidth", "50")));
        rotateApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("RotateApplySetWidth", "50")));
        rotateRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("RotateRemarkWidth", "120")));

        stockPoColumn.setPrefWidth(Double.valueOf(prefs.get("StockPoWidth", "120")));
        stockLotColumn.setPrefWidth(Double.valueOf(prefs.get("StockLotWidth", "50")));
        stockDcColumn.setPrefWidth(Double.valueOf(prefs.get("StockDcWidth", "120")));
        stockQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockQtyWidth", "63")));
        stockMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockMyQtyWidth", "50")));
        stockRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("StockRemarkWidth", "150")));

        purchaseGrDateColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseGrDateWidth", "120")));
        purchaseGrQtyColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseGrQtyWidth", "50")));
        purchaseMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseMyQtyWidth", "50")));
        purchaseApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseApplySetWidth", "50")));
        purchaseRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseRemarkWidth", "150")));

        noneStPurchasePoColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchasePoWidth", "120")));
        noneStPurchaseGrDateColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseGrDateWidth", "120")));
        noneStPurchaseGrQtyColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseGrQtyWidth", "50")));
        noneStPurchaseMyQtyColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseMyQtyWidth", "50")));
        noneStPurchaseApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseApplySetWidth", "50")));
        noneStPurchaseRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseRemarkWidth", "150")));
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
        excelParser.loadRotateExcel(file, 4, "F", "H", "R",
                "AP", "AQ", "AS");
        log.info("Process rotate table: Done");

        file = new File(pathStockFile.getText());
        excelParser.loadStockExcel(file, 2, "G", "E", "I",
                "J", "N", "P", "L", "M");
        log.info("Process stock table: Done");

        file = new File(pathPurchaseFile.getText());
        excelParser.loadPurchaseExcel(file, 4, "B", "D", "G",
                "F", "I", "K", "M", "J");
        log.info("Process purchase table: Done");


        final TreeItem<RotateItem> rotateRoot = new RecursiveTreeItem<>(
                collection.getRotateObsList(), RecursiveTreeObject::getChildren);
        rotateTreeTableView.setRoot(rotateRoot);
        rotateTreeTableView.requestFocus();
        rotateTreeTableView.getSelectionModel().select(0);

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
        initialPurchaseTable();
        initialNoneStockPurchaseTable();
    }

    private void initialRotateTable() {
        rotateTreeTableView.getSelectionModel().setCellSelectionEnabled(true);

        rotateKitColumn.setStyle("-fx-alignment: top-left; ");
        rotatePartColumn.setStyle("-fx-alignment: top-left; ");
        rotateRemarkColumn.setStyle("-fx-alignment: top-left; ");

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

        rotateRatioColumn.setCellValueFactory(cellData -> {
            if (rotateRatioColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().ratioProperty();
            } else {
                return rotateRatioColumn.getComputedValue(cellData);
            }
        });

        rotateApplySetColumn.setCellValueFactory(cellData -> {
            if (rotateApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return rotateApplySetColumn.getComputedValue(cellData);
            }
        });

        rotateRemarkColumn.setCellValueFactory(cellData -> {
            if (rotateRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return rotateRemarkColumn.getComputedValue(cellData);
            }
        });

        rotateKitColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<RotateItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        // set same kit with same background
                        if (item != null && !item.isEmpty() && item.compareTo(rotateSelectedKit) == 0) {
                            setStyle("-fx-alignment: top-left; -fx-background-color: #FFEB3B;");
                        } else {
                            setStyle("-fx-alignment: top-left; ");
                        }
                    }
                });
        rotateKitColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().kitNameProperty().set(cellData.getNewValue());
            rotateTreeTableView.requestFocus();
        });

        rotatePartColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<RotateItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void updateItem(String item, boolean empty) {
                        super.updateItem(item, empty);
                        // set same kit with same background
                        int rowIdx = getTreeTableRow().getIndex();
                        String kit = rotateKitColumn.getCellData(rowIdx);
                        if (item != null && !kit.isEmpty() && kit.compareTo(rotateSelectedKit) == 0) {
                            setStyle("-fx-alignment: top-left; -fx-background-color: #FFEB3B;");
                        } else {
                            setStyle("-fx-alignment: top-left; ");
                        }
                    }
                });
        rotatePartColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().partNumberProperty().set(cellData.getNewValue());
            rotateTreeTableView.requestFocus();
        });

        rotateMyQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<RotateItem, Number>(new IntegerTextFieldEditorBuilder()) {
            @Override
            public void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);
                // set same kit with same background
                int rowIdx = getTreeTableRow().getIndex();
                Number pmQty = rotatePmQtyColumn.getCellData(rowIdx);
                if (pmQty == null || item == null) {
                    setStyle("");
                } else if (pmQty.intValue() == item.intValue()) {
                    setStyle("-fx-background-color: #4CAF50;");
                } else if (pmQty.intValue() < item.intValue()) {
                    setStyle("-fx-background-color: #ef5350;");
                } else {
                    setStyle("");
                }

                setEditable(false);
            }
        });

        rotateApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        rotateApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            rotateTreeTableView.requestFocus();
        });

        rotateRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        rotateRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            rotateTreeTableView.requestFocus();
        });

        filterField.textProperty().addListener((o, oldVal, newVal) -> rotateTreeTableView.setPredicate(itemProp -> {
            final RotateItem item = itemProp.getValue();
            return item.kitNameProperty().get().contains(newVal) || item.partNumberProperty().get().contains(newVal);
        }));

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
                    setStyle("-fx-background-color: #607D8B;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void initialStockTable() {
        stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
        stockTreeTableView.getSelectionModel().setCellSelectionEnabled(true);
        //stockTreeTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        stockPoColumn.setStyle("-fx-alignment: top-left; ");
        stockLotColumn.setStyle("-fx-alignment: top-left; ");
        stockRemarkColumn.setStyle("-fx-alignment: top-left; ");

        stockPoColumn.setCellValueFactory(cellData -> {
            if (stockPoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return stockPoColumn.getComputedValue(cellData);
            }
        });

        stockLotColumn.setCellValueFactory(cellData -> {
            if (stockLotColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().lotProperty();
            } else {
                return stockLotColumn.getComputedValue(cellData);
            }
        });

        stockDcColumn.setCellValueFactory(cellData -> {
            if (stockDcColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().dcProperty();
            } else {
                return stockDcColumn.getComputedValue(cellData);
            }
        });

        stockQtyColumn.setCellValueFactory(cellData -> {
            if (stockQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().stockQtyProperty();
            } else {
                return stockQtyColumn.getComputedValue(cellData);
            }
        });

        stockMyQtyColumn.setCellValueFactory(cellData -> {
            if (stockMyQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().myQtyProperty();
            } else {
                return stockMyQtyColumn.getComputedValue(cellData);
            }
        });

        stockRemarkColumn.setCellValueFactory(cellData -> {
            if (stockRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return stockRemarkColumn.getComputedValue(cellData);
            }
        });

        stockMyQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        stockMyQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().myQtyProperty().set(cellData.getNewValue().intValue());
            stockTreeTableView.requestFocus();
        });

        stockRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        stockRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            stockTreeTableView.requestFocus();
        });

        stockTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedStockItem(oldValue, newValue)
        );

        stockTreeTableView.setRowFactory(tv -> new JFXTreeTableRow<StockItem>() {
            @Override
            public void updateItem(StockItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null) {
                    setStyle("");
                } else if (item.getPo().compareTo(stockSelectedPo) == 0) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void initialPurchaseTable() {
        purchaseTreeTableView.setPlaceholder(new Label("ZMM 找不到符合項目"));
        purchaseTreeTableView.getSelectionModel().setCellSelectionEnabled(true);
        //purchaseTreeTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        purchaseGrDateColumn.setCellValueFactory(cellData -> {
            if (purchaseGrDateColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grDateProperty();
            } else {
                return purchaseGrDateColumn.getComputedValue(cellData);
            }
        });

        purchaseGrQtyColumn.setCellValueFactory(cellData -> {
            if (purchaseGrQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grQtyProperty();
            } else {
                return purchaseGrQtyColumn.getComputedValue(cellData);
            }
        });

        purchaseMyQtyColumn.setCellValueFactory(cellData -> {
            if (purchaseMyQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().myQtyProperty();
            } else {
                return purchaseMyQtyColumn.getComputedValue(cellData);
            }
        });

        purchaseApplySetColumn.setCellValueFactory(cellData -> {
            if (purchaseApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return purchaseApplySetColumn.getComputedValue(cellData);
            }
        });

        purchaseRemarkColumn.setCellValueFactory(cellData -> {
            if (purchaseRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return purchaseRemarkColumn.getComputedValue(cellData);
            }
        });

        purchaseMyQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        purchaseMyQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().myQtyProperty().set(cellData.getNewValue().intValue());
            purchaseTreeTableView.requestFocus();
        });

        purchaseApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        purchaseApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            purchaseTreeTableView.requestFocus();
        });

        purchaseRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        purchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            purchaseTreeTableView.requestFocus();
        });

        purchaseTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedPurchaseItem(oldValue, newValue)
        );

        purchaseTreeTableView.setRowFactory(tv -> new JFXTreeTableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item != null && item.getPo().compareTo(purchaseSelectedPo) == 0) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });

    }

    private void initialNoneStockPurchaseTable() {
        noneStPurchaseTreeTableView.setPlaceholder(new Label("ZMM 找不到符合項目"));
        noneStPurchaseTreeTableView.getSelectionModel().setCellSelectionEnabled(true);
        //noneStPurchaseTreeTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        noneStPurchasePoColumn.setStyle("-fx-alignment: top-left; ");

        noneStPurchasePoColumn.setCellValueFactory(cellData -> {
            if (noneStPurchasePoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return noneStPurchasePoColumn.getComputedValue(cellData);
            }
        });

        noneStPurchaseGrDateColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseGrDateColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grDateProperty();
            } else {
                return noneStPurchaseGrDateColumn.getComputedValue(cellData);
            }
        });

        noneStPurchaseGrQtyColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseGrQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grQtyProperty();
            } else {
                return noneStPurchaseGrQtyColumn.getComputedValue(cellData);
            }
        });

        noneStPurchaseMyQtyColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseMyQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().myQtyProperty();
            } else {
                return noneStPurchaseMyQtyColumn.getComputedValue(cellData);
            }
        });

        noneStPurchaseApplySetColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return noneStPurchaseApplySetColumn.getComputedValue(cellData);
            }
        });

        noneStPurchaseRemarkColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return noneStPurchaseRemarkColumn.getComputedValue(cellData);
            }
        });


        noneStPurchaseMyQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        noneStPurchaseMyQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().myQtyProperty().set(cellData.getNewValue().intValue());
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        noneStPurchaseApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        noneStPurchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedNoneStockPurchaseItem(oldValue, newValue)
        );

        noneStPurchaseTreeTableView.setRowFactory(tv -> new JFXTreeTableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item != null && item.getPo().compareTo(noneStPurchaseSelectedPo) == 0) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void selectedRotateItem(TreeItem<RotateItem> oldValue, TreeItem<RotateItem> newValue) {

        if (newValue == null || !newValue.isLeaf()) {
            stockTreeTableView.setRoot(null);
            return;
        }

        RotateItem rotateItem = newValue.getValue();
        rotateSelectedKit = rotateItem.getKitName();
        rotateKitColumn.setVisible(false);
        rotateKitColumn.setVisible(true);

        ObservableList<StockItem> obsList = rotateItem.getStockItemObsList();
        TreeItem<StockItem> root = new RecursiveTreeItem<>(
                obsList, RecursiveTreeObject::getChildren);

        if (rotateItem.isDuplicate()) {
            root = null;
            stockTreeTableView.setPlaceholder(new Label("重複的項目!! 我無法幫妳了 orz...."));
        } else if (obsList.size() == 0) {
            stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
            root = null;
        }

        stockTreeTableView.setRoot(root);

        stockStQtyTotal.textProperty().bind(rotateItem.stockQtyTotalProperty().asString());
        stockGrQtyTotal.textProperty().bind(rotateItem.grQtyTotalProperty().asString());
        stockApQtyTotal.textProperty().bind(rotateItem.myQtyProperty().asString());


        ObservableList<PurchaseItem> noneStockObsList = rotateItem.getNoneStockPurchaseItemObsList();
        TreeItem<PurchaseItem> noneStockRoot = new RecursiveTreeItem<>(
                noneStockObsList, RecursiveTreeObject::getChildren);

        noneStPurchaseTreeTableView.setRoot(noneStockRoot);

        noneStPurchaseGrQtyTotal.textProperty().bind(rotateItem.noneStPurchaseGrQtyTotalProperty().asString());
        noneStPurchaseApQtyTotal.textProperty().bind(rotateItem.noneStPurchaseApQtyTotalProperty().asString());
        noneStPurchaseApSetTotal.textProperty().bind(rotateItem.noneStPurchaseApSetTotalProperty().asString());


        //stockTreeTableView.getSelectionModel().select(0);
    }


    private void selectedStockItem(TreeItem<StockItem> oldValue, TreeItem<StockItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            purchaseTreeTableView.setRoot(null);
            return;
        }

        StockItem stockItem = newValue.getValue();
        stockSelectedPo = stockItem.getPo();

        ObservableList<PurchaseItem> obsList = stockItem.getPurchaseItems();
        TreeItem<PurchaseItem> root = new RecursiveTreeItem<>(
                obsList, RecursiveTreeObject::getChildren);

        if (obsList.size() == 0) {
            root = null;
        }

        purchaseTreeTableView.setRoot(root);

        stockCurrPoStQtyTotal.textProperty().bind(stockItem.duplicatePoStockQtyTotalProperty().asString());
        stockCurrPoApQtyTotal.textProperty().bind(stockItem.duplicatePoMyQtyTotalProperty().asString());

        purchaseGrQtyTotal.textProperty().bind(stockItem.purchaseGrQtyTotalProperty().asString());
        purchaseApQtyTotal.textProperty().bind(stockItem.purchaseApQtyTotalProperty().asString());
        purchaseApSetTotal.textProperty().bind(stockItem.purchaseApSetTotalProperty().asString());
        //purchaseTreeTableView.getSelectionModel().select(0);
    }


    private void selectedPurchaseItem(TreeItem<PurchaseItem> oldValue, TreeItem<PurchaseItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            purchaseTreeTableView.setRoot(null);
            return;
        }

        PurchaseItem purchaseItem = newValue.getValue();
        purchaseSelectedPo = purchaseItem.getPo();
    }

    private void selectedNoneStockPurchaseItem(TreeItem<PurchaseItem> oldValue, TreeItem<PurchaseItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            purchaseTreeTableView.setRoot(null);
            return;
        }

        PurchaseItem purchaseItem = newValue.getValue();
        noneStPurchaseSelectedPo = purchaseItem.getPo();
    }

}
