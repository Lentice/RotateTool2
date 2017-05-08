package lahome.rotateTool.view;

import com.jfoenix.controls.*;
import com.jfoenix.controls.cells.editors.IntegerTextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.TextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.base.GenericEditableTreeTableCell;
import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.application.Platform;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import javafx.util.converter.DefaultStringConverter;
import javafx.util.converter.NumberStringConverter;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.ExcelSettings;
import lahome.rotateTool.Util.ExcelParser;
import lahome.rotateTool.Util.ExcelSaver;
import lahome.rotateTool.module.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.util.prefs.Preferences;

public class RootController {
    private static final Logger log = LogManager.getLogger(RootController.class.getName());

    @FXML
    private JFXTabPane mainTabPane;

    @FXML
    private JFXButton exportToExcelButton;

    @FXML
    private JFXButton startEditButton;

    @FXML
    private Tab editTab;

    @FXML
    private SplitPane editTabHSlitPane;

    @FXML
    private SplitPane editTabVLeftSlitPane;

    @FXML
    private SplitPane editTabVRightSlitPane;

    @FXML
    private JFXTextField pathAgingReport;

    @FXML
    private JFXTextField pathStockFile;

    @FXML
    private JFXTextField pathPurchaseFile;

    @FXML
    private TextField settingRotateFirstRow;

    @FXML
    private TextField settingRotateKitNameCol;

    @FXML
    private TextField settingRotatePartNumCol;

    @FXML
    private TextField settingRotatePmQtyCol;

    @FXML
    private TextField settingRotateApQtyCol;

    @FXML
    private TextField settingRotateRatioCol;

    @FXML
    private TextField settingRotateApSetCol;

    @FXML
    private TextField settingRotateRemarkCol;

    @FXML
    private TextField settingStockFirstRow;

    @FXML
    private TextField settingStockKitNameCol;

    @FXML
    private TextField settingStockPartNumCol;

    @FXML
    private TextField settingStockPoCol;

    @FXML
    private TextField settingStockStockQtyCol;

    @FXML
    private TextField settingStockApQtyCol;

    @FXML
    private TextField settingStockRemarkCol;

    @FXML
    private TextField settingStockLotCol;

    @FXML
    private TextField settingStockDcCol;

    @FXML
    private TextField settingPurchaseFirstRow;

    @FXML
    private TextField settingPurchaseKitNameCol;

    @FXML
    private TextField settingPurchasePartNumCol;

    @FXML
    private TextField settingPurchasePoCol;

    @FXML
    private TextField settingPurchaseGrDateCol;

    @FXML
    private TextField settingPurchaseGrQtyCol;

    @FXML
    private TextField settingPurchaseApQtyCol;

    @FXML
    private TextField settingPurchaseApSetCol;

    @FXML
    private TextField settingPurchaseRemarkCol;

    @FXML
    private ImageView filterIcon;

    @FXML
    private JFXTextField rotateFilterField;

    @FXML
    private JFXTreeTableView<RotateItem> rotateTableView;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateKitColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotatePartColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateNoColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotatePmQtyColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotateApQtyColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateRatioColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, Number> rotateApplySetColumn;

    @FXML
    private JFXTreeTableColumn<RotateItem, String> rotateRemarkColumn;

    @FXML
    private Label rotatePartCount;

    @FXML
    private Label rotateCurrKitZmmSetTotal;

    @FXML
    private ComboBox<String> stockPartNumCombo;

    @FXML
    private TableView<StockItem> stockTableView;

    @FXML
    private TableColumn<StockItem, String> stockNoColumn;

    @FXML
    private TableColumn<StockItem, String> stockPartNumColumn;

    @FXML
    private TableColumn<StockItem, String> stockPoColumn;

    @FXML
    private TableColumn<StockItem, String> stockLotColumn;

    @FXML
    private TableColumn<StockItem, String> stockDcColumn;

    @FXML
    private TableColumn<StockItem, String> stockGrDateColumn;

    @FXML
    private TableColumn<StockItem, Number> stockStQtyColumn;

    @FXML
    private TableColumn<StockItem, Number> stockApQtyColumn;

    @FXML
    private TableColumn<StockItem, String> stockRemarkColumn;

    @FXML
    private Label stockPmQty;

    @FXML
    private Label stockStQtyTotal;

    @FXML
    private Label stockApQtyTotal;

    @FXML
    private Label stockCurrPoStQtyTotal;

    @FXML
    private Label stockCurrPoApQtyTotal;

    @FXML
    private Label purchaseTableTitle;

    @FXML
    private JFXToggleButton purchaseShowAllToggle;

    @FXML
    private TableView<PurchaseItem> purchaseTableView;

    @FXML
    private TableColumn<PurchaseItem, String> purchaseNoColumn;

    @FXML
    private TableColumn<PurchaseItem, String> purchasePartNumColumn;

    @FXML
    private TableColumn<PurchaseItem, String> purchasePoColumn;

    @FXML
    private TableColumn<PurchaseItem, String> purchaseGrDateColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> purchaseGrQtyColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> purchaseApQtyColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> purchaseApplySetColumn;

    @FXML
    private TableColumn<PurchaseItem, String> purchaseRemarkColumn;

    @FXML
    private Label purchaseGrQtyTotal;

    @FXML
    private Label purchaseApQtyTotal;

    @FXML
    private Label purchaseApSetTotal;

    @FXML
    private TableView<PurchaseItem> noneStPurchaseTableView;

    @FXML
    private TableColumn<PurchaseItem, String> noneStPurchaseNoColumn;

    @FXML
    private TableColumn<PurchaseItem, String> noneStPurchasePartNumColumn;

    @FXML
    private TableColumn<PurchaseItem, String> noneStPurchasePoColumn;

    @FXML
    private TableColumn<PurchaseItem, String> noneStPurchaseGrDateColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> noneStPurchaseGrQtyColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> noneStPurchaseApQtyColumn;

    @FXML
    private TableColumn<PurchaseItem, Number> noneStPurchaseApplySetColumn;

    @FXML
    private TableColumn<PurchaseItem, String> noneStPurchaseRemarkColumn;

    @FXML
    private Label noneStPurchaseGrQtyTotal;

    @FXML
    private Label noneStPurchaseApQtyTotal;

    @FXML
    private Label noneStPurchaseApSetTotal;


    private Main main;
    private Stage primaryStage;
    private RotateCollection collection;
    private ExcelParser excelParser;

    private File latestDir = new File(System.getProperty("user.dir"));
    private String rotateSelectedKit = "";
    private String stockSelectedPart = "";
    private String stockSelectedPo = "";
    private String purchaseSelectedPo = "";
    private String noneStPurchaseSelectedPo = "";
    private boolean showAllParts = true;

    private final String ALL_PARTS = "All";
    private RotateItem currentRotateItem;
    private StockItem currentStockItem;

    private static final Object excelHandleLock = new Object();

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

        prefs.put("settingRotateFirstRow", settingRotateFirstRow.getText());
        prefs.put("settingRotateKitNameCol", settingRotateKitNameCol.getText());
        prefs.put("settingRotatePartNumCol", settingRotatePartNumCol.getText());
        prefs.put("settingRotatePmQtyCol", settingRotatePmQtyCol.getText());
        prefs.put("settingRotateApQtyCol", settingRotateApQtyCol.getText());
        prefs.put("settingRotateRatioCol", settingRotateRatioCol.getText());
        prefs.put("settingRotateApSetCol", settingRotateApSetCol.getText());
        prefs.put("settingRotateRemarkCol", settingRotateRemarkCol.getText());

        prefs.put("settingStockFirstRow", settingStockFirstRow.getText());
        prefs.put("settingStockKitNameCol", settingStockKitNameCol.getText());
        prefs.put("settingStockPartNumCol", settingStockPartNumCol.getText());
        prefs.put("settingStockPoCol", settingStockPoCol.getText());
        prefs.put("settingStockStockQtyCol", settingStockStockQtyCol.getText());
        prefs.put("settingStockApQtyCol", settingStockApQtyCol.getText());
        prefs.put("settingStockRemarkCol", settingStockRemarkCol.getText());
        prefs.put("settingStockLotCol", settingStockLotCol.getText());
        prefs.put("settingStockDcCol", settingStockDcCol.getText());

        prefs.put("settingPurchaseFirstRow", settingPurchaseFirstRow.getText());
        prefs.put("settingPurchaseKitNameCol", settingPurchaseKitNameCol.getText());
        prefs.put("settingPurchasePartNumCol", settingPurchasePartNumCol.getText());
        prefs.put("settingPurchasePoCol", settingPurchasePoCol.getText());
        prefs.put("settingPurchaseGrDateCol", settingPurchaseGrDateCol.getText());
        prefs.put("settingPurchaseGrQtyCol", settingPurchaseGrQtyCol.getText());
        prefs.put("settingPurchaseApQtyCol", settingPurchaseApQtyCol.getText());
        prefs.put("settingPurchaseApSetCol", settingPurchaseApSetCol.getText());
        prefs.put("settingPurchaseRemarkCol", settingPurchaseRemarkCol.getText());

        prefs.put("EditTabHDiv", String.valueOf(editTabHSlitPane.getDividerPositions()[0]));
        prefs.put("EditTabVLeftDiv", String.valueOf(editTabVLeftSlitPane.getDividerPositions()[0]));
        prefs.put("EditTabVRightDiv", String.valueOf(editTabVRightSlitPane.getDividerPositions()[0]));

        prefs.put("RotateKitWidth", String.valueOf(rotateKitColumn.getWidth()));
        prefs.put("RotatePartWidth", String.valueOf(rotatePartColumn.getWidth()));
        prefs.put("RotateNoWidth", String.valueOf(rotateNoColumn.getWidth()));
        prefs.put("RotatePmQtyWidth", String.valueOf(rotatePmQtyColumn.getWidth()));
        prefs.put("RotateApQtyWidth", String.valueOf(rotateApQtyColumn.getWidth()));
        prefs.put("RotateRatioWidth", String.valueOf(rotateRatioColumn.getWidth()));
        prefs.put("RotateApplySetWidth", String.valueOf(rotateApplySetColumn.getWidth()));
        prefs.put("RotateRemarkWidth", String.valueOf(rotateRemarkColumn.getWidth()));

        prefs.put("StockNoWidth", String.valueOf(stockNoColumn.getWidth()));
        prefs.put("StockPartNumWidth", String.valueOf(stockPartNumColumn.getWidth()));
        prefs.put("StockPoWidth", String.valueOf(stockPoColumn.getWidth()));
        prefs.put("StockLotWidth", String.valueOf(stockLotColumn.getWidth()));
        prefs.put("StockDcWidth", String.valueOf(stockDcColumn.getWidth()));
        prefs.put("StockGrDateWidth", String.valueOf(stockGrDateColumn.getWidth()));
        prefs.put("StockQtyWidth", String.valueOf(stockStQtyColumn.getWidth()));
        prefs.put("StockApQtyWidth", String.valueOf(stockApQtyColumn.getWidth()));
        prefs.put("StockRemarkWidth", String.valueOf(stockRemarkColumn.getWidth()));

        prefs.put("PurchaseNoColumnWidth", String.valueOf(purchaseNoColumn.getWidth()));
        prefs.put("PurchasePartNumWidth", String.valueOf(purchasePartNumColumn.getWidth()));
        prefs.put("PurchasePoColumnWidth", String.valueOf(purchasePoColumn.getWidth()));
        prefs.put("PurchaseGrDateWidth", String.valueOf(purchaseGrDateColumn.getWidth()));
        prefs.put("PurchaseGrQtyWidth", String.valueOf(purchaseGrQtyColumn.getWidth()));
        prefs.put("PurchaseApQtyWidth", String.valueOf(purchaseApQtyColumn.getWidth()));
        prefs.put("PurchaseApplySetWidth", String.valueOf(purchaseApplySetColumn.getWidth()));
        prefs.put("PurchaseRemarkWidth", String.valueOf(purchaseRemarkColumn.getWidth()));

        prefs.put("NoneStPurchaseNoColumnWidth", String.valueOf(noneStPurchaseNoColumn.getWidth()));
        prefs.put("NoneStPurchasePartNumWidth", String.valueOf(noneStPurchasePartNumColumn.getWidth()));
        prefs.put("NoneStPurchasePoWidth", String.valueOf(noneStPurchasePoColumn.getWidth()));
        prefs.put("NoneStPurchaseGrDateWidth", String.valueOf(noneStPurchaseGrDateColumn.getWidth()));
        prefs.put("NoneStPurchaseGrQtyWidth", String.valueOf(noneStPurchaseGrQtyColumn.getWidth()));
        prefs.put("NoneStPurchaseApQtyWidth", String.valueOf(noneStPurchaseApQtyColumn.getWidth()));
        prefs.put("NoneStPurchaseApplySetWidth", String.valueOf(noneStPurchaseApplySetColumn.getWidth()));
        prefs.put("NoneStPurchaseRemarkWidth", String.valueOf(noneStPurchaseRemarkColumn.getWidth()));
    }

    public void loadSetting() {
        Preferences prefs = Preferences.userNodeForPackage(Main.class);

        latestDir = new File(prefs.get("LatestDir", System.getProperty("user.dir")));
        if (!latestDir.exists())
            latestDir = new File(System.getProperty("user.dir"));

        pathAgingReport.setText(prefs.get("AgingPath", ""));
        pathStockFile.setText(prefs.get("StockPath", ""));
        pathPurchaseFile.setText(prefs.get("PurchasePath", ""));

        settingRotateFirstRow.setText(prefs.get("settingRotateFirstRow", "4"));
        settingRotateKitNameCol.setText(prefs.get("settingRotateKitNameCol", "F"));
        settingRotatePartNumCol.setText(prefs.get("settingRotatePartNumCol", "H"));
        settingRotatePmQtyCol.setText(prefs.get("settingRotatePmQtyCol", "R"));
        settingRotateApQtyCol.setText(prefs.get("settingRotateApQtyCol", "AO"));
        settingRotateRatioCol.setText(prefs.get("settingRotateRatioCol", "AP"));
        settingRotateApSetCol.setText(prefs.get("settingRotateApSetCol", "AQ"));
        settingRotateRemarkCol.setText(prefs.get("settingRotateRemarkCol", "AS"));

        settingStockFirstRow.setText(prefs.get("settingStockFirstRow", "2"));
        settingStockKitNameCol.setText(prefs.get("settingStockKitNameCol", "G"));
        settingStockPartNumCol.setText(prefs.get("settingStockPartNumCol", "E"));
        settingStockPoCol.setText(prefs.get("settingStockPoCol", "I"));
        settingStockStockQtyCol.setText(prefs.get("settingStockStockQtyCol", "J"));
        settingStockApQtyCol.setText(prefs.get("settingStockApQtyCol", "L"));
        settingStockRemarkCol.setText(prefs.get("settingStockRemarkCol", "M"));
        settingStockLotCol.setText(prefs.get("settingStockLotCol", "N"));
        settingStockDcCol.setText(prefs.get("settingStockDcCol", "P"));

        settingPurchaseFirstRow.setText(prefs.get("settingPurchaseFirstRow", "4"));
        settingPurchaseKitNameCol.setText(prefs.get("settingPurchaseKitNameCol", "B"));
        settingPurchasePartNumCol.setText(prefs.get("settingPurchasePartNumCol", "D"));
        settingPurchasePoCol.setText(prefs.get("settingPurchasePoCol", "G"));
        settingPurchaseGrDateCol.setText(prefs.get("settingPurchaseGrDateCol", "F"));
        settingPurchaseGrQtyCol.setText(prefs.get("settingPurchaseGrQtyCol", "I"));
        settingPurchaseApQtyCol.setText(prefs.get("settingPurchaseApQtyCol", "K"));
        settingPurchaseApSetCol.setText(prefs.get("settingPurchaseApSetCol", "M"));
        settingPurchaseRemarkCol.setText(prefs.get("settingPurchaseRemarkCol", "J"));

        editTabHSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabHDiv", "0.41")));
        editTabVLeftSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabVLeftDiv", "0.75")));
        editTabVRightSlitPane.setDividerPosition(0, Double.valueOf(prefs.get("EditTabVRightDiv", "0.75")));

        rotateKitColumn.setPrefWidth(Double.valueOf(prefs.get("RotateKitWidth", "140")));
        rotatePartColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePartWidth", "140")));
        rotateNoColumn.setPrefWidth(Double.valueOf(prefs.get("RotateNoWidth", "30")));
        rotatePmQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotatePmQtyWidth", "50")));
        rotateApQtyColumn.setPrefWidth(Double.valueOf(prefs.get("RotateApQtyWidth", "50")));
        rotateRatioColumn.setPrefWidth(Double.valueOf(prefs.get("RotateRatioWidth", "50")));
        rotateApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("RotateApplySetWidth", "50")));
        rotateRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("RotateRemarkWidth", "120")));

        stockNoColumn.setPrefWidth(Double.valueOf(prefs.get("StockNoWidth", "30")));
        stockPartNumColumn.setPrefWidth(Double.valueOf(prefs.get("StockPartNumWidth", "120")));
        stockPoColumn.setPrefWidth(Double.valueOf(prefs.get("StockPoWidth", "120")));
        stockLotColumn.setPrefWidth(Double.valueOf(prefs.get("StockLotWidth", "120")));
        stockDcColumn.setPrefWidth(Double.valueOf(prefs.get("StockDcWidth", "50")));
        stockGrDateColumn.setPrefWidth(Double.valueOf(prefs.get("StockGrDateWidth", "50")));
        stockStQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockQtyWidth", "63")));
        stockApQtyColumn.setPrefWidth(Double.valueOf(prefs.get("StockApQtyWidth", "50")));
        stockRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("StockRemarkWidth", "150")));

        purchaseNoColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseNoColumnWidth", "30")));
        purchasePartNumColumn.setPrefWidth(Double.valueOf(prefs.get("PurchasePartNumWidth", "120")));
        purchasePoColumn.setPrefWidth(Double.valueOf(prefs.get("PurchasePoColumnWidth", "120")));
        purchaseGrDateColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseGrDateWidth", "120")));
        purchaseGrQtyColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseGrQtyWidth", "50")));
        purchaseApQtyColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseApQtyWidth", "50")));
        purchaseApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseApplySetWidth", "50")));
        purchaseRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("PurchaseRemarkWidth", "150")));

        noneStPurchaseNoColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseNoColumnWidth", "30")));
        noneStPurchasePartNumColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchasePartNumWidth", "120")));
        noneStPurchasePoColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchasePoWidth", "120")));
        noneStPurchaseGrDateColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseGrDateWidth", "120")));
        noneStPurchaseGrQtyColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseGrQtyWidth", "50")));
        noneStPurchaseApQtyColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseApQtyWidth", "50")));
        noneStPurchaseApplySetColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseApplySetWidth", "50")));
        noneStPurchaseRemarkColumn.setPrefWidth(Double.valueOf(prefs.get("NoneStPurchaseRemarkWidth", "150")));
    }

    public void setMainApp(Main main) {
        this.main = main;
        this.primaryStage = main.getPrimaryStage();
        this.collection = main.getCollection();
        excelParser = new ExcelParser(collection);
        initialHotkeys();
    }


    private void initialHotkeys() {
        Scene scene = primaryStage.getScene();

        scene.setOnKeyPressed(keyEvent -> {
            Node focusNode = scene.getFocusOwner();

            if (keyEvent.getCode() == KeyCode.F12) {
                int row = rotateTableView.getSelectionModel().getSelectedIndex();
                row++;
                if (row < rotateTableView.getCurrentItemsCount()) {
                    rotateTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTableView.scrollToColumn(rotateApplySetColumn);
                }
                focusNode.requestFocus();

                //Stop letting it do anything else
                keyEvent.consume();
            }

            if (keyEvent.getCode() == KeyCode.F11) {

                int row = rotateTableView.getSelectionModel().getSelectedIndex();
                row--;
                if (row >= 0) {
                    rotateTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTableView.scrollToColumn(rotateApplySetColumn);
                }
                focusNode.requestFocus();

                //Stop letting it do anything else
                keyEvent.consume();
            }
        });
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

        synchronized (excelHandleLock) {
            ExcelSettings.setRotateConfig(
                    pathAgingReport.getText(),
                    Integer.valueOf(settingRotateFirstRow.getText()),
                    settingRotateKitNameCol.getText(),
                    settingRotatePartNumCol.getText(),
                    settingRotatePmQtyCol.getText(),
                    settingRotateApQtyCol.getText(),
                    settingRotateRatioCol.getText(),
                    settingRotateApSetCol.getText(),
                    settingRotateRemarkCol.getText());

            ExcelSettings.setStockConfig(
                    pathStockFile.getText(),
                    Integer.valueOf(settingStockFirstRow.getText()),
                    settingStockKitNameCol.getText(),
                    settingStockPartNumCol.getText(),
                    settingStockPoCol.getText(),
                    settingStockStockQtyCol.getText(),
                    settingStockLotCol.getText(),
                    settingStockDcCol.getText(),
                    settingStockApQtyCol.getText(),
                    settingStockRemarkCol.getText());

            ExcelSettings.setPurchaseConfig(
                    pathPurchaseFile.getText(),
                    Integer.valueOf(settingPurchaseFirstRow.getText()),
                    settingPurchaseKitNameCol.getText(),
                    settingPurchasePartNumCol.getText(),
                    settingPurchasePoCol.getText(),
                    settingPurchaseGrDateCol.getText(),
                    settingPurchaseGrQtyCol.getText(),
                    settingPurchaseApQtyCol.getText(),
                    settingPurchaseApSetCol.getText(),
                    settingPurchaseRemarkCol.getText());

            Platform.runLater(() -> startEditButton.setDisable(true));
            importExcel();
        }
    }

    private void importExcel() {
        new Thread(() -> {
            synchronized (excelHandleLock) {
                excelParser.loadRotateExcel();
                log.info("Process rotate table: Done");

                excelParser.loadStockExcel();
                log.info("Process stock table: Done");

                excelParser.loadPurchaseExcel();
                log.info("Process purchase table: Done");

                StringBuilder errorLog = new StringBuilder("");
                int failCount = DataChecker.doCheck(collection, errorLog);
                if (failCount > 0) {
                    Platform.runLater(() -> {
                        Alert alert = new Alert(Alert.AlertType.WARNING);

                        alert.setTitle("Warning");
                        alert.setHeaderText(String.format("Total %d invalid data were found", failCount));
                        alert.setContentText("Error logs are listed below:");

                        TextArea textArea = new TextArea(errorLog.toString());
                        textArea.setEditable(false);
                        textArea.setWrapText(false);

                        textArea.setMaxWidth(Double.MAX_VALUE);
                        textArea.setMaxHeight(Double.MAX_VALUE);
                        GridPane.setVgrow(textArea, Priority.ALWAYS);
                        GridPane.setHgrow(textArea, Priority.ALWAYS);

                        GridPane expContent = new GridPane();
                        expContent.setMaxWidth(Double.MAX_VALUE);
                        expContent.add(textArea, 0, 0);

                        // Set expandable Exception into the dialog pane.
                        alert.getDialogPane().setExpandableContent(expContent);
                        alert.getDialogPane().setExpanded(true);

                        alert.getDialogPane().setPrefWidth(600);
                        alert.showAndWait();
                    });
                }

                Platform.runLater(() -> {
                    mainTabPane.getSelectionModel().select(editTab);
                    updateRotateTable();
                    exportToExcelButton.setVisible(true);
                });

                startEditButton.setDisable(false);
            }
        }).start();

    }

    @FXML
    void handleSaveFiles() {
        log.info("Save File clicked");
        ExcelSaver excelSaver = new ExcelSaver(collection);

        excelSaver.saveRotateToExcel();
        log.info("Rotate file saved");
        excelSaver.saveStockToExcel();
        log.info("Stock file saved");
        excelSaver.savePurchaseToExcel();
        log.info("Purchase file saved");
    }

    @FXML
    private void initialize() {

        filterIcon.setImage(new Image("file:resources/images/active-search-2-48.png"));
        exportToExcelButton.setVisible(false);
        initialRotateTable();
        initialStockTable();
        initialPurchaseTable();
        initialNoneStockPurchaseTable();
    }

    private void initialRotateTable() {
        rotateTableView.getSelectionModel().setCellSelectionEnabled(true);

        rotateKitColumn.setStyle("-fx-alignment: top-left; ");
        rotatePartColumn.setStyle("-fx-alignment: top-left; ");
        rotateNoColumn.setStyle("-fx-alignment: top-left; ");
        rotateRemarkColumn.setStyle("-fx-alignment: top-left; ");

        rotateKitColumn.setCellValueFactory(cellData -> {
            if (rotateKitColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().kitNameProperty();
            } else {
                return rotateKitColumn.getComputedValue(cellData);
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

                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        rotatePartColumn.setCellValueFactory(cellData -> {
            if (rotatePartColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().partNumberProperty();
            } else {
                return rotatePartColumn.getComputedValue(cellData);
            }
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

                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        rotateNoColumn.setCellValueFactory(cellData -> {
            if (rotateNoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().serialNoProperty();
            } else {
                return rotateNoColumn.getComputedValue(cellData);
            }
        });
        rotateNoColumn.setCellFactory(cellData ->
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

                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        rotatePmQtyColumn.setCellValueFactory(cellData -> {
            if (rotatePmQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().pmQtyProperty();
            } else {
                return rotatePmQtyColumn.getComputedValue(cellData);
            }
        });
        rotatePmQtyColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<RotateItem, Number>(new IntegerTextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(Number newValue) {
                    }
                });

        rotateApQtyColumn.setCellValueFactory(cellData -> {
            if (rotateApQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().stockApplyQtyTotalProperty();
            } else {
                return rotateApQtyColumn.getComputedValue(cellData);
            }
        });
        rotateApQtyColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<RotateItem, Number>(new IntegerTextFieldEditorBuilder()) {
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
                    }

                    @Override
                    public void commitEdit(Number newValue) {
                    }
                });

        rotateRatioColumn.setCellValueFactory(cellData -> {
            if (rotateRatioColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().ratioProperty();
            } else {
                return rotateRatioColumn.getComputedValue(cellData);
            }
        });
        rotateRatioColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<RotateItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        rotateApplySetColumn.setCellValueFactory(cellData -> {
            if (rotateApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return rotateApplySetColumn.getComputedValue(cellData);
            }
        });
        rotateApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        rotateApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            rotateTableView.requestFocus();
        });

        rotateRemarkColumn.setCellValueFactory(cellData -> {
            if (rotateRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return rotateRemarkColumn.getComputedValue(cellData);
            }
        });
        rotateRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        rotateRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            rotateTableView.requestFocus();
        });

        rotateFilterField.textProperty().addListener((o, oldVal, newVal) -> rotateTableView.setPredicate(itemProp -> {
            final RotateItem item = itemProp.getValue();
            return item.kitNameProperty().get().contains(newVal) || item.partNumberProperty().get().contains(newVal);
        }));

        rotateTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedRotateItem(newValue));

        rotateTableView.setRowFactory(tv -> new JFXTreeTableRow<RotateItem>() {
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

        stockPartNumCombo.autosize();
        stockPartNumCombo.setOnAction(e -> stockPartNumComboChanged());
    }

    private void initialStockTable() {
        stockTableView.setPlaceholder(new Label("庫存找不到符合項目"));
        stockTableView.getSelectionModel().setCellSelectionEnabled(true);
        stockTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        stockTableView.setEditable(true);

        Callback<TableColumn<StockItem, String>, TableCell<StockItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<StockItem, String>(new DefaultStringConverter()) {
            @Override
            public void commitEdit(String newValue) {
            }
        };

        Callback<TableColumn<StockItem, Number>, TableCell<StockItem, Number>> readOnlyNumberCell
                = p -> new TextFieldTableCell<StockItem, Number>(new NumberStringConverter()) {
            @Override
            public void commitEdit(Number newValue) {
            }
        };

        stockNoColumn.getStyleClass().add("my-table-column-part-serial");
        stockNoColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().serialNoProperty());
        stockNoColumn.setCellFactory(readOnlyStringCell);

        stockPartNumColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().partNumberProperty());
        stockPartNumColumn.setCellFactory(readOnlyStringCell);

        stockPoColumn.setCellValueFactory(cellData -> cellData.getValue().poProperty());
        stockPoColumn.setCellFactory(readOnlyStringCell);

        stockLotColumn.setCellValueFactory(cellData -> cellData.getValue().lotProperty());
        stockLotColumn.setCellFactory(readOnlyStringCell);

        stockDcColumn.setCellValueFactory(cellData -> cellData.getValue().dcProperty());
        stockDcColumn.setCellFactory(readOnlyStringCell);


        stockGrDateColumn.setCellValueFactory(cellData -> cellData.getValue().earliestGrDateProperty());
        stockGrDateColumn.setCellFactory(readOnlyStringCell);

        stockStQtyColumn.getStyleClass().add("my-table-column-number");
        stockStQtyColumn.setCellValueFactory(cellData -> cellData.getValue().stockQtyProperty());
        stockStQtyColumn.setCellFactory(readOnlyNumberCell);

        stockApQtyColumn.getStyleClass().add("my-table-column-number");
        stockApQtyColumn.setCellValueFactory(cellData -> cellData.getValue().applyQtyProperty());
        stockApQtyColumn.setCellFactory(TextFieldTableCell.forTableColumn(new NumberStringConverter()));
        stockApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updateStockTableTotal();
            stockTableView.requestFocus();
        });

        stockRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        stockRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        stockRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            stockTableView.requestFocus();
        });

        stockTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedStockItem(newValue));

        stockTableView.setRowFactory(param -> new TableRow<StockItem>() {
            @Override
            public void updateItem(StockItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
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
        purchaseTableView.setPlaceholder(new Label("ZMM 找不到符合項目"));
        purchaseTableView.getSelectionModel().setCellSelectionEnabled(true);
        purchaseTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        purchaseTableView.setEditable(true);

        Callback<TableColumn<PurchaseItem, String>, TableCell<PurchaseItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<PurchaseItem, String>(new DefaultStringConverter()) {
            @Override
            public void commitEdit(String newValue) {
            }
        };

        Callback<TableColumn<PurchaseItem, Number>, TableCell<PurchaseItem, Number>> readOnlyNumberCell
                = p -> new TextFieldTableCell<PurchaseItem, Number>(new NumberStringConverter()) {
            @Override
            public void commitEdit(Number newValue) {
            }
        };

        purchaseNoColumn.getStyleClass().add("my-table-column-part-serial");
        purchaseNoColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().serialNoProperty());
        purchaseNoColumn.setCellFactory(readOnlyStringCell);

        purchasePartNumColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().partNumberProperty());
        purchasePartNumColumn.setCellFactory(readOnlyStringCell);

        purchasePoColumn.setCellValueFactory(cellData -> cellData.getValue().poProperty());
        purchasePoColumn.setCellFactory(readOnlyStringCell);

        purchaseGrDateColumn.setCellValueFactory(cellData -> cellData.getValue().grDateProperty());
        purchaseGrDateColumn.setCellFactory(readOnlyStringCell);

        purchaseGrQtyColumn.getStyleClass().add("my-table-column-number");
        purchaseGrQtyColumn.setCellValueFactory(cellData -> cellData.getValue().grQtyProperty());
        purchaseGrQtyColumn.setCellFactory(readOnlyNumberCell);

        purchaseApQtyColumn.getStyleClass().add("my-table-column-number");
        purchaseApQtyColumn.setCellValueFactory(cellData -> cellData.getValue().applyQtyProperty());
        purchaseApQtyColumn.setCellFactory(TextFieldTableCell.forTableColumn(new NumberStringConverter()));
        purchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updatePurchaseTableTotal();
            stockTableView.requestFocus();
        });

        purchaseApplySetColumn.getStyleClass().add("my-table-column-number");
        purchaseApplySetColumn.setCellValueFactory(cellData -> cellData.getValue().applySetProperty());
        purchaseApplySetColumn.setCellFactory(TextFieldTableCell.forTableColumn(new NumberStringConverter()));
        purchaseApplySetColumn.setCellFactory(readOnlyNumberCell);

        purchaseRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        purchaseRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        purchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            stockTableView.requestFocus();
        });

        purchaseTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedPurchaseItem(newValue));

        purchaseShowAllToggle.setOnAction(a -> purchaseShowAllChanged());

        purchaseTableView.setRowFactory(param -> new TableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.getPo().compareTo(purchaseSelectedPo) == 0) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void initialNoneStockPurchaseTable() {
        noneStPurchaseTableView.setPlaceholder(new Label("ZMM 找不到符合項目"));
        noneStPurchaseTableView.getSelectionModel().setCellSelectionEnabled(true);
        noneStPurchaseTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        purchaseTableView.setEditable(true);

        Callback<TableColumn<PurchaseItem, String>, TableCell<PurchaseItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<PurchaseItem, String>(new DefaultStringConverter()) {
            @Override
            public void commitEdit(String newValue) {
            }
        };

        Callback<TableColumn<PurchaseItem, Number>, TableCell<PurchaseItem, Number>> readOnlyNumberCell
                = p -> new TextFieldTableCell<PurchaseItem, Number>(new NumberStringConverter()) {
            @Override
            public void commitEdit(Number newValue) {
            }
        };

        noneStPurchaseNoColumn.getStyleClass().add("my-table-column-part-serial");
        noneStPurchaseNoColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().serialNoProperty());
        noneStPurchaseNoColumn.setCellFactory(readOnlyStringCell);

        noneStPurchasePartNumColumn.setCellValueFactory(cellData -> cellData.getValue().getRotateItem().partNumberProperty());
        noneStPurchasePartNumColumn.setCellFactory(readOnlyStringCell);

        noneStPurchasePoColumn.setCellValueFactory(cellData -> cellData.getValue().poProperty());
        noneStPurchasePoColumn.setCellFactory(readOnlyStringCell);

        noneStPurchaseGrDateColumn.setCellValueFactory(cellData -> cellData.getValue().grDateProperty());
        noneStPurchaseGrDateColumn.setCellFactory(readOnlyStringCell);

        noneStPurchaseGrQtyColumn.setCellValueFactory(cellData -> cellData.getValue().grQtyProperty());
        noneStPurchaseGrQtyColumn.setCellFactory(readOnlyNumberCell);

        noneStPurchaseApQtyColumn.setCellValueFactory(cellData -> cellData.getValue().applyQtyProperty());
        noneStPurchaseApQtyColumn.setCellFactory(TextFieldTableCell.forTableColumn(new NumberStringConverter()));
        noneStPurchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updateNoneStPurchaseTableTotal();
            stockTableView.requestFocus();
        });

        noneStPurchaseApplySetColumn.setCellValueFactory(cellData -> cellData.getValue().applySetProperty());
        noneStPurchaseApplySetColumn.setCellFactory(TextFieldTableCell.forTableColumn(new NumberStringConverter()));
        noneStPurchaseApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applySetProperty().set(cellData.getNewValue().intValue());
            updateNoneStPurchaseTableTotal();
            stockTableView.requestFocus();
        });

        noneStPurchaseRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        noneStPurchaseRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        noneStPurchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            stockTableView.requestFocus();
        });

        noneStPurchaseTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedNoneStockPurchaseItem(newValue)
        );

        noneStPurchaseTableView.setRowFactory(param -> new TableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.getPo().compareTo(noneStPurchaseSelectedPo) == 0) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void selectedRotateItem(TreeItem<RotateItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            stockTableView.getItems().clear();
            return;
        }

        rotateKitColumn.setVisible(false);
        rotateKitColumn.setVisible(true);

        RotateItem rotateItem = newValue.getValue();
        boolean kitUpdate = currentRotateItem == null || !rotateItem.isKit() ||
                rotateItem.getKitName().compareTo(currentRotateItem.getKitName()) != 0;
        if (!kitUpdate)
            return;

        currentRotateItem = rotateItem;
        rotateSelectedKit = rotateItem.getKitName();

        if (rotateItem.isKit()) {
            stockPartNumCombo.getItems().clear();
            for (RotateItem part : rotateItem.getKitNode().getPartsList()) {
                stockPartNumCombo.getItems().add(part.getPartNumber());
            }

            showAllParts = true;
            stockPartNumCombo.getItems().add(ALL_PARTS);
            stockPartNumCombo.setValue(ALL_PARTS);

        } else {
            showAllParts = false;
            stockPartNumCombo.getItems().clear();
            stockPartNumCombo.getItems().add(rotateItem.getPartNumber());
            stockPartNumCombo.setValue(rotateItem.getPartNumber());
        }

        updateRotateTableTotal();

        updateStockTable(rotateItem);
        updateNoneStPurchaseTable(rotateItem);
    }

    private void selectedStockItem(StockItem newStockItem) {
        if (newStockItem == null) {
            purchaseTableView.getItems().clear();
            return;
        }

        stockPoColumn.setVisible(false);
        stockPoColumn.setVisible(true);

        currentStockItem = newStockItem;
        stockSelectedPo = newStockItem.getPo();
        stockSelectedPart = newStockItem.getRotateItem().getPartNumber();

        updatePurchaseTable(newStockItem);
    }

    private void selectedPurchaseItem(PurchaseItem newPurchaseItem) {
        if (newPurchaseItem == null)
            return;

        purchaseSelectedPo = newPurchaseItem.getPo();
    }

    private void selectedNoneStockPurchaseItem(PurchaseItem newPurchaseItem) {
        if (newPurchaseItem == null) {
            return;
        }

        noneStPurchaseSelectedPo = newPurchaseItem.getPo();
    }

    private void stockPartNumComboChanged() {
        if (currentRotateItem == null)
            return;

        if (stockPartNumCombo.getItems().size() == 0)
            return;

        String name = stockPartNumCombo.getSelectionModel().getSelectedItem().trim();
        if (name.isEmpty())
            return;

        RotateItem rotateItem = currentRotateItem;
        if (currentRotateItem.isKit())
            rotateItem = currentRotateItem.getKitNode().getPart(name);

        if (rotateItem != null) {
            showAllParts = false;
            updateStockTable(rotateItem);
            updateNoneStPurchaseTable(rotateItem);
        } else {
            if (!showAllParts) {
                showAllParts = true;
                updateStockTable(currentRotateItem);
                updateNoneStPurchaseTable(currentRotateItem);
            }
        }

    }

    private void purchaseShowAllChanged() {
        if (currentStockItem != null) {
            purchasePoColumn.setVisible(false);
            purchasePoColumn.setVisible(true);
            updatePurchaseTable(currentStockItem);
        }
    }

    private void updateRotateTable() {
        final TreeItem<RotateItem> rotateRoot = new RecursiveTreeItem<>(
                collection.getRotateObsList(), RecursiveTreeObject::getChildren);

        rotateTableView.setRoot(rotateRoot);
        if (rotateTableView.getCurrentItemsCount() > 0) {
            rotateTableView.getSelectionModel().select(0, rotatePmQtyColumn);
        }

        rotateTableView.requestFocus();
        rotatePartCount.setText(String.valueOf(collection.getRotateObsList().size()));
    }

    private void updateStockTable(RotateItem rotateItem) {

        if (rotateItem.isDuplicate()) {
            stockTableView.setPlaceholder(new Label("重複的項目!! 我無法幫妳了 orz...."));
            stockTableView.getItems().clear();
        } else {
            ObservableList<StockItem> obsList;
            if (rotateItem.isKit() && showAllParts) {
                obsList = rotateItem.getKitNode().getStockItemList();
            } else {
                obsList = rotateItem.getStockItemObsList();
            }

            if (obsList.size() > 0) {
                stockTableView.setItems(obsList);
                stockTableView.getSelectionModel().select(0, stockApQtyColumn);
            } else {
                stockTableView.setPlaceholder(new Label("庫存找不到符合項目"));
            }
        }

        updateStockTableTotal();

        if (rotateItem.isKit() && showAllParts) {
            stockPmQty.setText(String.valueOf(rotateItem.getKitNode().pmQtyProperty()));
        } else {
            stockPmQty.setText(String.valueOf(rotateItem.getPmQty()));
        }

    }

    private void updatePurchaseTable(StockItem stockItem) {


        if (stockItem.getRotateItem().isDuplicate()) {
            purchaseTableView.getItems().clear();
        } else {
            ObservableList<PurchaseItem> obsList;
            if (purchaseShowAllToggle.isSelected()) {
                if (stockItem.getRotateItem().isKit() && showAllParts) {
                    obsList = stockItem.getRotateItem().getKitNode().getPurchaseItemList();
                } else {
                    obsList = stockItem.getRotateItem().getPurchaseItemList();
                }
            } else {
                if (stockItem.getRotateItem().isKit() && showAllParts) {
                    obsList = stockItem.getRotateItem().getKitNode().getPurchaseItemListWithPo(stockItem.getPo());
                } else {
                    obsList = stockItem.getPurchaseItems();
                }
            }

            if (obsList.size() > 0) {
                purchaseTableView.setItems(obsList);
                purchaseTableView.getSelectionModel().select(0, purchaseApQtyColumn);
            }
        }

        updatePurchaseTableTotal();

        if (purchaseShowAllToggle.isSelected()) {
            purchaseTableTitle.setText("ZMM (All PO)");
        } else {
            purchaseTableTitle.setText("ZMM (PO:" + stockItem.getPo() + ")");
        }

    }

    private void updateNoneStPurchaseTable(RotateItem rotateItem) {

        if (rotateItem.isDuplicate()) {
            noneStPurchaseTableView.getItems().clear();
        } else {
            ObservableList<PurchaseItem> obsList;
            if (rotateItem.isKit() && showAllParts) {
                obsList = rotateItem.getKitNode().getNoneStPurchaseItemList();
            } else {
                obsList = rotateItem.getNoneStockPurchaseItemObsList();
            }

            if (obsList.size() > 0) {
                noneStPurchaseTableView.setItems(obsList);
                noneStPurchaseTableView.getSelectionModel().select(0, noneStPurchaseApQtyColumn);
            }
        }

        updateNoneStPurchaseTableTotal();
    }

    private void updateRotateTableTotal() {
        int applySetTotal = 0;

        if (currentRotateItem.isKit()) {
            for (PurchaseItem purchaseItem : currentRotateItem.getKitNode().getStAndNoneStPurchaseItemList()) {
                applySetTotal += purchaseItem.getApplySet();
            }
        } else {
            for (StockItem stockItem : currentRotateItem.getStockItemObsList()) {
                if (stockItem.isMainStockItem()) {
                    for (PurchaseItem purchaseItem : stockItem.getPurchaseItems()) {
                        applySetTotal += purchaseItem.getApplySet();
                    }
                }
            }
        }

        rotateCurrKitZmmSetTotal.setText(String.valueOf(applySetTotal));
    }

    private void updateStockTableTotal() {
        int stQtyTotal = 0;
        int applyQtyTotal = 0;

        int currentPoStQtyTotal = 0;
        int currentPoApplyQtyTotal = 0;

        ObservableList<StockItem> list = stockTableView.getItems();
        if (list != null) {
            for (StockItem stockItem : list) {
                stQtyTotal += stockItem.getStockQty();
                applyQtyTotal += stockItem.getApplyQty();

                if (stockSelectedPart.equals(stockItem.getRotateItem().getPartNumber()) &&
                        stockSelectedPo.equals(stockItem.getPo())) {
                    currentPoStQtyTotal += stockItem.getStockQty();
                    currentPoApplyQtyTotal += stockItem.getApplyQty();
                }
            }
        }

        stockStQtyTotal.setText(String.valueOf(stQtyTotal));
        stockApQtyTotal.setText(String.valueOf(applyQtyTotal));

        stockCurrPoStQtyTotal.setText(String.valueOf(currentPoStQtyTotal));
        stockCurrPoApQtyTotal.setText(String.valueOf(currentPoApplyQtyTotal));
    }

    private void updatePurchaseTableTotal() {
        int grQtyTotal = 0;
        int applyQtyTotal = 0;
        int applySetTotal = 0;

        ObservableList<PurchaseItem> list = purchaseTableView.getItems();
        if (list != null) {
            for (PurchaseItem purchaseItem : list) {

                grQtyTotal += purchaseItem.getGrQty();
                applyQtyTotal += purchaseItem.getApplyQty();
                applySetTotal += purchaseItem.getApplySet();
            }
        }

        purchaseGrQtyTotal.setText(String.valueOf(grQtyTotal));
        purchaseApQtyTotal.setText(String.valueOf(applyQtyTotal));
        purchaseApSetTotal.setText(String.valueOf(applySetTotal));

    }

    private void updateNoneStPurchaseTableTotal() {
        int grQtyTotal = 0;
        int applyQtyTotal = 0;
        int applySetTotal = 0;

        ObservableList<PurchaseItem> list = noneStPurchaseTableView.getItems();
        if (list != null) {
            for (PurchaseItem purchaseItem : list) {
                grQtyTotal += purchaseItem.getGrQty();
                applyQtyTotal += purchaseItem.getApplyQty();
                applySetTotal += purchaseItem.getApplySet();
            }
        }

        noneStPurchaseGrQtyTotal.setText(String.valueOf(grQtyTotal));
        noneStPurchaseApQtyTotal.setText(String.valueOf(applyQtyTotal));
        noneStPurchaseApSetTotal.setText(String.valueOf(applySetTotal));

    }

}
