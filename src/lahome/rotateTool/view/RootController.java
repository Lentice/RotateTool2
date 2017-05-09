package lahome.rotateTool.view;

import com.jfoenix.controls.*;
import javafx.application.Platform;
import javafx.beans.property.DoubleProperty;
import javafx.beans.property.SimpleDoubleProperty;
import javafx.beans.value.ChangeListener;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.collections.transformation.SortedList;
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
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import javafx.util.converter.DefaultStringConverter;
import javafx.util.converter.NumberStringConverter;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.Excel.ExcelSettings;
import lahome.rotateTool.Util.Excel.ExcelParser;
import lahome.rotateTool.Util.Excel.ExcelSaver;
import lahome.rotateTool.Util.TableUtils;
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
    private VBox processStatus;

    @FXML
    private ProgressIndicator rotateProgress;

    @FXML
    private ProgressIndicator stockProgress;

    @FXML
    private ProgressIndicator purchaseProgress;

    @FXML
    private TextField settingRotateFirstRow;

    @FXML
    private TextField settingRotateKitNameCol;

    @FXML
    private TextField settingRotatePartNumCol;

    @FXML
    private TextField settingRotateBacklogCol;

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
    private TableView<RotateItem> rotateTableView;

    @FXML
    private TableColumn<RotateItem, String> rotateKitColumn;

    @FXML
    private TableColumn<RotateItem, String> rotatePartColumn;

    @FXML
    private TableColumn<RotateItem, String> rotateNoColumn;

    @FXML
    private TableColumn<RotateItem, String> rotateBacklogColumn;

    @FXML
    private TableColumn<RotateItem, Number> rotatePmQtyColumn;

    @FXML
    private TableColumn<RotateItem, Number> rotateApQtyColumn;

    @FXML
    private TableColumn<RotateItem, String> rotateRatioColumn;

    @FXML
    private TableColumn<RotateItem, Number> rotateApplySetColumn;

    @FXML
    private TableColumn<RotateItem, String> rotateRemarkColumn;

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

    private File latestDir = new File(System.getProperty("user.dir"));
    private String rotateSelectedKit = "";
    private String rotateSelectedPart = "";
    private String stockSelectedPart = "";
    private String stockSelectedPo = "";
    private String purchaseSelectedPo = "";
    private String noneStPurchaseSelectedPo = "";
    private boolean showAllParts = true;
    private final String ALL_PARTS = "All";
    private RotateItem currentRotateItem;
    private StockItem currentStockItem;
    private boolean rotateFilterListenerAdded = false;

    private ObservableList<RotateItem> emptyRotateList = FXCollections.observableArrayList();
    private ObservableList<StockItem> emptyStockList = FXCollections.observableArrayList();
    private ObservableList<PurchaseItem> emptyPurchaseList = FXCollections.observableArrayList();

    private DoubleProperty rotateProgressProperty = new SimpleDoubleProperty(0);
    private DoubleProperty stockProgressProperty = new SimpleDoubleProperty(0);
    private DoubleProperty purchaseProgressProperty = new SimpleDoubleProperty(0);


    public RootController() {
    }

    @FXML
    private void initialize() {

        filterIcon.setImage(new Image(Main.class.getResourceAsStream("/images/active-search-2-48.png")));
        exportToExcelButton.setVisible(false);
        initialRotateTable();
        initialStockTable();
        initialPurchaseTable();
        initialNoneStockPurchaseTable();
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
        prefs.put("settingRotateBacklogCol", settingRotateBacklogCol.getText());
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
        prefs.put("RotateBacklogWidth", String.valueOf(rotateBacklogColumn.getWidth()));
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
        settingRotateBacklogCol.setText(prefs.get("settingRotateBacklogCol", "Q"));
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
        rotateBacklogColumn.setPrefWidth(Double.valueOf(prefs.get("RotateBacklogWidth", "60")));
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
        initialHotkeys();

        rotateProgressProperty.addListener((observable, oldValue, newValue) ->
                Platform.runLater(() -> rotateProgress.setProgress(newValue.doubleValue())));
        stockProgressProperty.addListener((observable, oldValue, newValue) ->
                Platform.runLater(() -> stockProgress.setProgress(newValue.doubleValue())));
        purchaseProgressProperty.addListener((observable, oldValue, newValue) ->
                Platform.runLater(() -> purchaseProgress.setProgress(newValue.doubleValue())));

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
        startEditButton.setDisable(true);

        rotateProgress.setProgress(0);
        stockProgress.setProgress(0);
        purchaseProgress.setProgress(0);

        configExcelSettings();
        importExcel();
    }

    @FXML
    void handleSaveFiles() {
        log.info("Save File clicked");

        exportToExcelButton.setDisable(true);
        rotateProgress.setProgress(0);
        stockProgress.setProgress(0);
        purchaseProgress.setProgress(0);

        new Thread(() -> {

            ExcelSaver excelSaver = new ExcelSaver(collection);

            excelSaver.saveRotateToExcel(rotateProgressProperty);
            log.info("Rotate file saved");
            excelSaver.saveStockToExcel(stockProgressProperty);
            log.info("Stock file saved");
            excelSaver.savePurchaseToExcel(purchaseProgressProperty);
            log.info("Purchase file saved");

            Platform.runLater(() -> exportToExcelButton.setDisable(false));
        }).start();
    }

    private void configExcelSettings() {
        ExcelSettings.setRotateConfig(
                pathAgingReport.getText(),
                Integer.valueOf(settingRotateFirstRow.getText()),
                settingRotateKitNameCol.getText(),
                settingRotatePartNumCol.getText(),
                settingRotateBacklogCol.getText(),
                settingRotatePmQtyCol.getText(),
                settingRotateApQtyCol.getText(),
                settingRotateRatioCol.getText(),
                settingRotateApSetCol.getText(),
                settingRotateRemarkCol.getText()
        );

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
                settingStockRemarkCol.getText()
        );

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
                settingPurchaseRemarkCol.getText()
        );
    }

    private void importExcel() {

        new Thread(() -> {
            ExcelParser excelParser = new ExcelParser(collection);
            StringBuilder errorLog = new StringBuilder("");

            excelParser.loadRotateExcel(rotateProgressProperty);
            log.info("Process rotate table: Done");

            excelParser.loadStockExcel(stockProgressProperty);
            log.info("Process stock table: Done");

            excelParser.loadPurchaseExcel(purchaseProgressProperty);
            log.info("Process purchase table: Done");

            int failCount = DataChecker.doCheck(collection, errorLog);
            if (failCount > 0) {
                showAlertWarning(
                        "Warning",
                        String.format("Total %d invalid data were found", failCount),
                        "Error logs are listed below:",
                        errorLog.toString()
                );
            }

            try {
                Thread.sleep(500);  // waiting user to see all item were completed.
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            Platform.runLater(() -> {
                mainTabPane.getSelectionModel().select(editTab);
                updateRotateTable();
                exportToExcelButton.setVisible(true);
            });

        }).start();
    }

    private void showAlertWarning(String title, String header, String content, String detail) {
        Platform.runLater(() -> {
            Alert alert = new Alert(Alert.AlertType.WARNING);

            alert.setTitle(title);
            alert.setHeaderText(header);
            alert.setContentText(content);

            TextArea textArea = new TextArea(detail);
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

    private void initialHotkeys() {
        Scene scene = primaryStage.getScene();

        scene.setOnKeyPressed(keyEvent -> {
            Node focusNode = scene.getFocusOwner();

            if (keyEvent.getCode() == KeyCode.F12) {
                keyEvent.consume();

                int row = rotateTableView.getSelectionModel().getSelectedIndex();
                row++;
                if (row < rotateTableView.getItems().size()) {
                    rotateTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTableView.scrollToColumn(rotateApplySetColumn);
                }
                focusNode.requestFocus();
            }

            if (keyEvent.getCode() == KeyCode.F11) {
                keyEvent.consume();

                int row = rotateTableView.getSelectionModel().getSelectedIndex();
                row--;
                if (row >= 0) {
                    rotateTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTableView.scrollToColumn(rotateApplySetColumn);
                }
                focusNode.requestFocus();
            }

            if (keyEvent.getCode() == KeyCode.F1) {
                keyEvent.consume();

                int listCount = stockPartNumCombo.getItems().size();
                if (listCount <= 1) {
                    return;
                }

                // get current item
                int index = 0;
                for (String item : stockPartNumCombo.getItems()) {
                    if (item.equals(stockPartNumCombo.getValue())) {
                        break;
                    }
                    index++;
                }

                // Select next item
                if (++index >= listCount) {
                    stockPartNumCombo.setValue(stockPartNumCombo.getItems().get(0));
                } else {
                    stockPartNumCombo.setValue(stockPartNumCombo.getItems().get(index));
                }

                focusNode.requestFocus();
            }
        });
    }

    private void initialRotateTable() {
        rotateTableView.getSelectionModel().setCellSelectionEnabled(true);
        rotateTableView.setEditable(true);
        TableUtils.installCopyPasteHandler(rotateTableView);

        Callback<TableColumn<RotateItem, String>, TableCell<RotateItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<RotateItem, String>(new DefaultStringConverter()) {
            @Override
            public void commitEdit(String newValue) {
            }
        };

        Callback<TableColumn<RotateItem, Number>, TableCell<RotateItem, Number>> readOnlyNumberCell
                = p -> new TextFieldTableCell<RotateItem, Number>(new NumberStringConverter()) {
            @Override
            public void commitEdit(Number newValue) {
            }
        };

        rotateKitColumn.setCellValueFactory(cellData -> cellData.getValue().kitNameProperty());
        rotateKitColumn.setCellFactory(cellData -> new TextFieldTableCell<RotateItem, String>() {
            @Override
            public void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);

                // set same kit with same background
                String basicStyle = "-fx-alignment: center-left; ";
                RotateItem rotateItem = (RotateItem) this.getTableRow().getItem();

                if (rotateItem == null || rotateItem.isDuplicate() || !rotateItem.isKit()) {
                    setStyle(basicStyle);
                } else if (rotateItem.isKit() && rotateItem.getKitName().equals(rotateSelectedKit)) {
                    setStyle(basicStyle + "-fx-background-color: #FFEB3B;");
                } else {
                    setStyle(basicStyle);
                }
            }

            @Override
            public void commitEdit(String newValue) {
            }
        });

        rotatePartColumn.setCellValueFactory(cellData -> cellData.getValue().partNumberProperty());
        rotatePartColumn.setCellFactory(cellData -> new TextFieldTableCell<RotateItem, String>() {
            @Override
            public void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);

                // set same kit with same background
                String basicStyle = "-fx-alignment: center-left; ";
                RotateItem rotateItem = (RotateItem) this.getTableRow().getItem();

                if (rotateItem == null || rotateItem.isDuplicate()) {
                    setStyle(basicStyle);
                } else if (!rotateItem.isKit() && rotateItem.getPartNumber().equals(rotateSelectedPart)) {
                    setStyle(basicStyle + "-fx-background-color: #FFEB3B;");
                } else if (rotateItem.isKit() && rotateItem.getKitName().equals(rotateSelectedKit)) {
                    setStyle(basicStyle + "-fx-background-color: #FFEB3B;");
                } else {
                    setStyle(basicStyle);
                }
            }

            @Override
            public void commitEdit(String newValue) {
            }
        });

        rotateNoColumn.setCellValueFactory(cellData -> cellData.getValue().serialNoProperty());
        rotateNoColumn.setCellFactory(cellData -> new TextFieldTableCell<RotateItem, String>() {
            @Override
            public void updateItem(String item, boolean empty) {
                super.updateItem(item, empty);

                // set same kit with same background
                String basicStyle = "-fx-alignment: center-left; ";
                RotateItem rotateItem = (RotateItem) this.getTableRow().getItem();
                if (rotateItem == null || rotateItem.isDuplicate()) {
                    setStyle(basicStyle);
                } else if (!rotateItem.isKit() && rotateItem.getPartNumber().equals(rotateSelectedPart)) {
                    setStyle(basicStyle + "-fx-background-color: #FFEB3B;");
                } else if (rotateItem.isKit() && rotateItem.getKitName().equals(rotateSelectedKit)) {
                    setStyle(basicStyle + "-fx-background-color: #FFEB3B;");
                } else {
                    setStyle(basicStyle);
                }
            }

            @Override
            public void commitEdit(String newValue) {
            }
        });

        rotateBacklogColumn.getStyleClass().add("my-table-column-number");
        rotateBacklogColumn.setCellValueFactory(cellData -> cellData.getValue().backlogProperty());
        rotateBacklogColumn.setCellFactory(readOnlyStringCell);

        rotatePmQtyColumn.setCellValueFactory(cellData -> cellData.getValue().pmQtyProperty());
        rotatePmQtyColumn.setCellFactory(readOnlyNumberCell);

        rotateApQtyColumn.setCellValueFactory(cellData -> cellData.getValue().stockApplyQtyTotalProperty());
        rotateApQtyColumn.setCellFactory(cellData -> new TextFieldTableCell<RotateItem, Number>(new NumberStringConverter()) {
            @Override
            public void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);

                String basicStyle = "-fx-alignment: center-right; ";
                RotateItem rotateItem = (RotateItem) this.getTableRow().getItem();
                if (item == null || empty || rotateItem == null) {
                    setStyle(basicStyle);
                    return;
                }

                int pmQty = rotateItem.getPmQty();
                if (item.intValue() > pmQty) {
                    setStyle(basicStyle + "-fx-background-color: #ef5350;"); // red
                } else if (rotateItem.getStockApplyQtyTotal() != rotateItem.getPurchasesApplyQtyTotal()) {
                    setStyle(basicStyle + "-fx-background-color: #BA68C8;"); // purple
                } else if ((rotateItem.getStockApplyQtyTotal() % rotateItem.getRatio()) != 0) {
                    setStyle(basicStyle + "-fx-background-color: #FFAB91;"); // orange
                } else if (item.intValue() == pmQty) {
                    setStyle(basicStyle + "-fx-background-color: #4CAF50;"); // green
                } else {
                    setStyle(basicStyle);
                }
            }

            @Override
            public void commitEdit(Number newValue) {
            }
        });

        rotateRatioColumn.getStyleClass().add("my-table-column-number");
        rotateRatioColumn.setCellValueFactory(cellData -> cellData.getValue().ratioProperty());
        rotateRatioColumn.setCellFactory(readOnlyStringCell);

        rotateApplySetColumn.getStyleClass().add("my-table-column-number");
        rotateApplySetColumn.setCellValueFactory(cellData -> cellData.getValue().applySetProperty());
        rotateApplySetColumn.setCellFactory(readOnlyNumberCell);

        rotateRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        rotateRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        rotateRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            refreshAllTable();
            rotateTableView.requestFocus();
        });

        rotateTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> handleSelectedRotateItem(newValue));

        rotateTableView.setRowFactory(tv -> new TableRow<RotateItem>() {
            @Override
            public void updateItem(RotateItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.isDuplicate()) {
                    setStyle("-fx-background-color: #607D8B;");
                } else {
                    setStyle("");
                }
            }
        });

        stockPartNumCombo.autosize();
        stockPartNumCombo.setOnAction(e -> handleStockPartNumComboChanged());
    }

    private void initialStockTable() {
        stockTableView.setPlaceholder(new Label("庫存找不到符合項目"));
        stockTableView.getSelectionModel().setCellSelectionEnabled(true);
        stockTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
        stockTableView.setEditable(true);
        TableUtils.installCopyPasteHandler(stockTableView);

        Callback<TableColumn<StockItem, String>, TableCell<StockItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<StockItem, String>() {
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
        stockApQtyColumn.setCellFactory(cellData -> new TextFieldTableCell<StockItem, Number>(new NumberStringConverter()) {
            @Override
            public void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);

                String basicStyle = "";
                StockItem stockItem = (StockItem) this.getTableRow().getItem();
                if (item == null || empty || stockItem == null) {
                    setStyle(basicStyle);
                    return;
                }


                if (item.intValue() > stockItem.getStockQty()) {
                    setStyle(basicStyle + "-fx-background-color: #ef5350;"); // red
                //} else if ((item.intValue() % stockItem.getRotateItem().getRatio()) != 0) {
                //    setStyle(basicStyle + "-fx-background-color: #BA68C8;"); // purple
                } else {
                    setStyle(basicStyle);
                }
            }
        });
        stockApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updateStockTableTotal();
            refreshAllTable();
            stockTableView.requestFocus();
        });

        stockRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        stockRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        stockRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            refreshAllTable();
            stockTableView.requestFocus();
        });

        stockTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> handleSelectedStockItem(newValue));

        stockTableView.setRowFactory(param -> new TableRow<StockItem>() {
            @Override
            public void updateItem(StockItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.getPo().equals(stockSelectedPo)) {
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
        TableUtils.installCopyPasteHandler(purchaseTableView);

        Callback<TableColumn<PurchaseItem, String>, TableCell<PurchaseItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<PurchaseItem, String>() {
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
        purchaseApQtyColumn.setCellFactory(cellData -> new TextFieldTableCell<PurchaseItem, Number>(new NumberStringConverter()) {
            @Override
            public void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);

                String basicStyle = "";
                PurchaseItem purchaseItem = (PurchaseItem) this.getTableRow().getItem();
                if (item == null || empty || purchaseItem == null) {
                    setStyle(basicStyle);
                    return;
                }

                if (item.intValue() > purchaseItem.getGrQty()) {
                    setStyle(basicStyle + "-fx-background-color: #ef5350;"); // red
                //} else if ((item.intValue() % purchaseItem.getRotateItem().getRatio()) != 0) {
                //    setStyle(basicStyle + "-fx-background-color: #BA68C8;"); // purple
                } else {
                    setStyle(basicStyle);
                }
            }
        });
        purchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updatePurchaseTableTotal();
            refreshAllTable();
            purchaseTableView.requestFocus();
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
            refreshAllTable();
            purchaseTableView.requestFocus();
        });

        purchaseTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> handleSelectedPurchaseItem(newValue));

        purchaseShowAllToggle.setOnAction(a -> handlePurchaseShowAllChanged());

        purchaseTableView.setRowFactory(param -> new TableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.getPo().equals(purchaseSelectedPo)) {
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
        noneStPurchaseTableView.setEditable(true);
        TableUtils.installCopyPasteHandler(noneStPurchaseTableView);

        Callback<TableColumn<PurchaseItem, String>, TableCell<PurchaseItem, String>> readOnlyStringCell
                = p -> new TextFieldTableCell<PurchaseItem, String>() {
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

        noneStPurchaseGrQtyColumn.getStyleClass().add("my-table-column-number");
        noneStPurchaseGrQtyColumn.setCellValueFactory(cellData -> cellData.getValue().grQtyProperty());
        noneStPurchaseGrQtyColumn.setCellFactory(readOnlyNumberCell);

        noneStPurchaseApQtyColumn.getStyleClass().add("my-table-column-number");
        noneStPurchaseApQtyColumn.setCellValueFactory(cellData -> cellData.getValue().applyQtyProperty());
        noneStPurchaseApQtyColumn.setCellFactory(cellData -> new TextFieldTableCell<PurchaseItem, Number>(new NumberStringConverter()) {
            @Override
            public void updateItem(Number item, boolean empty) {
                super.updateItem(item, empty);

                String basicStyle = "";
                PurchaseItem purchaseItem = (PurchaseItem) this.getTableRow().getItem();
                if (item == null || empty || purchaseItem == null) {
                    setStyle(basicStyle);
                    return;
                }

                if (item.intValue() > purchaseItem.getGrQty()) {
                    setStyle(basicStyle + "-fx-background-color: #ef5350;"); // red
                //} else if ((item.intValue() % purchaseItem.getRotateItem().getRatio()) != 0) {
                //    setStyle(basicStyle + "-fx-background-color: #BA68C8;"); // purple
                } else {
                    setStyle(basicStyle);
                }
            }
        });
        noneStPurchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .applyQtyProperty().set(cellData.getNewValue().intValue());
            updateNoneStPurchaseTableTotal();
            refreshAllTable();
            noneStPurchaseTableView.requestFocus();
        });

        noneStPurchaseApplySetColumn.getStyleClass().add("my-table-column-number");
        noneStPurchaseApplySetColumn.setCellValueFactory(cellData -> cellData.getValue().applySetProperty());
        noneStPurchaseApplySetColumn.setCellFactory(readOnlyNumberCell);

        noneStPurchaseRemarkColumn.setCellValueFactory(cellData -> cellData.getValue().remarkProperty());
        noneStPurchaseRemarkColumn.setCellFactory(TextFieldTableCell.forTableColumn());
        noneStPurchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTableView().getItems().get(cellData.getTablePosition().getRow())
                    .remarkProperty().set(cellData.getNewValue());
            refreshAllTable();
            noneStPurchaseTableView.requestFocus();
        });

        noneStPurchaseTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> handleSelectedNoneStockPurchaseItem(newValue)
        );

        noneStPurchaseTableView.setRowFactory(param -> new TableRow<PurchaseItem>() {
            @Override
            public void updateItem(PurchaseItem item, boolean empty) {
                super.updateItem(item, empty);
                if (item == null || empty) {
                    setStyle("");
                } else if (item.getPo().equals(noneStPurchaseSelectedPo)) {
                    setStyle("-fx-background-color: #FFF176;");
                } else {
                    setStyle("");
                }
            }
        });
    }

    private void refreshAllTable() {
        rotateKitColumn.setVisible(false);
        rotateKitColumn.setVisible(true);

        stockPoColumn.setVisible(false);
        stockPoColumn.setVisible(true);

        purchasePoColumn.setVisible(false);
        purchasePoColumn.setVisible(true);

        noneStPurchasePoColumn.setVisible(false);
        noneStPurchasePoColumn.setVisible(true);
    }

    private void handleSelectedRotateItem(RotateItem newRotateItem) {
        if (newRotateItem == null) {
            stockTableView.setItems(emptyStockList);
            noneStPurchaseTableView.setItems(emptyPurchaseList);
            return;
        }

        rotateKitColumn.setVisible(false);
        rotateKitColumn.setVisible(true);

        boolean kitUpdate = currentRotateItem == null || !newRotateItem.isKit() ||
                !newRotateItem.getKitName().equals(currentRotateItem.getKitName());
        if (!kitUpdate)
            return;

        currentRotateItem = newRotateItem;
        rotateSelectedKit = newRotateItem.getKitName();
        rotateSelectedPart = newRotateItem.getPartNumber();

        if (newRotateItem.isKit()) {
            stockPartNumCombo.getItems().clear();
            for (RotateItem part : newRotateItem.getKitNode().getPartsList()) {
                stockPartNumCombo.getItems().add(part.getPartNumber());
            }

            showAllParts = true;
            stockPartNumCombo.getItems().add(ALL_PARTS);
            stockPartNumCombo.setValue(ALL_PARTS);

        } else {
            showAllParts = false;
            stockPartNumCombo.getItems().clear();
            stockPartNumCombo.getItems().add(newRotateItem.getPartNumber());
            stockPartNumCombo.setValue(newRotateItem.getPartNumber());
        }


        updateStockTable(newRotateItem);
        updateNoneStPurchaseTable(newRotateItem);

        updateRotateTableTotal();
    }

    private void handleSelectedStockItem(StockItem newStockItem) {
        if (newStockItem == null) {
            purchaseTableView.setItems(emptyPurchaseList);
            return;
        }

        stockPoColumn.setVisible(false);
        stockPoColumn.setVisible(true);

        currentStockItem = newStockItem;
        stockSelectedPo = newStockItem.getPo();
        stockSelectedPart = newStockItem.getRotateItem().getPartNumber();

        updatePurchaseTable(newStockItem);
        updateStockTableTotal();
    }

    private void handleSelectedPurchaseItem(PurchaseItem newPurchaseItem) {
        if (newPurchaseItem == null)
            return;

        purchasePoColumn.setVisible(false);
        purchasePoColumn.setVisible(true);

        purchaseSelectedPo = newPurchaseItem.getPo();
        updatePurchaseTableTotal();
    }

    private void handleSelectedNoneStockPurchaseItem(PurchaseItem newPurchaseItem) {
        if (newPurchaseItem == null) {
            return;
        }

        noneStPurchasePoColumn.setVisible(false);
        noneStPurchasePoColumn.setVisible(true);

        noneStPurchaseSelectedPo = newPurchaseItem.getPo();
        updateNoneStPurchaseTableTotal();
    }

    private void handleStockPartNumComboChanged() {
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

    private void handlePurchaseShowAllChanged() {
        if (currentStockItem != null) {
            purchasePoColumn.setVisible(false);
            purchasePoColumn.setVisible(true);
            updatePurchaseTable(currentStockItem);
        }
    }

    private void updateRotateTable() {

        rotateTableView.setItems(emptyRotateList);
        ObservableList<RotateItem> obsList = collection.getRotateObsList();
        FilteredList<RotateItem> filteredData = new FilteredList<>(obsList, p -> true);

        if (!rotateFilterListenerAdded) {
            rotateFilterListenerAdded = true;

            ChangeListener listener = (ChangeListener<String>) (ov, oldVal, newVal) -> filteredData.setPredicate(rotateItem -> {
                // If filter text is empty, display all persons.
                if (newVal == null || newVal.isEmpty()) {
                    return true;
                }

                // Compare first name and last name of every person with filter text.
                String lowerCaseFilter = newVal.toLowerCase();

                if (rotateItem.getKitName().toLowerCase().contains(lowerCaseFilter)) {
                    return true; // Filter matches first name.
                } else if (rotateItem.getPartNumber().toLowerCase().contains(lowerCaseFilter)) {
                    return true; // Filter matches last name.
                }
                return false; // Does not match.
            });

            rotateFilterField.textProperty().addListener(listener);
        }

        // 3. Wrap the FilteredList in a SortedList.
        SortedList<RotateItem> sortedData = new SortedList<>(filteredData);

        // 4. Bind the SortedList comparator to the TableView comparator.
        sortedData.comparatorProperty().bind(rotateTableView.comparatorProperty());

        // 5. Add sorted (and filtered) data to the table.
        rotateTableView.setItems(sortedData);
        if (obsList.size() > 0) {
            rotateTableView.getSelectionModel().select(0, rotateApplySetColumn);
        }

        rotateTableView.requestFocus();
        rotatePartCount.setText(String.valueOf(obsList.size()));
    }

    private void updateStockTable(RotateItem rotateItem) {

        stockTableView.setItems(emptyStockList);
        if (rotateItem.isDuplicate()) {
            stockTableView.setPlaceholder(new Label("重複的項目!! 我無法幫妳了 orz...."));
        } else {
            stockTableView.setPlaceholder(new Label("庫存找不到符合項目"));

            ObservableList<StockItem> obsList;
            if (rotateItem.isKit() && showAllParts) {
                obsList = rotateItem.getKitNode().getStockItemList();
            } else {
                obsList = rotateItem.getStockItemObsList();
            }

            stockTableView.setItems(obsList);
            if (obsList.size() > 0) {
                stockTableView.getSelectionModel().select(0, stockApQtyColumn);
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

        purchaseTableView.setItems(emptyPurchaseList);
        if (!stockItem.getRotateItem().isDuplicate()) {
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

            purchaseTableView.setItems(obsList);
            if (obsList.size() > 0) {
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

        noneStPurchaseTableView.setItems(emptyPurchaseList);
        if (!rotateItem.isDuplicate()) {
            ObservableList<PurchaseItem> obsList;
            if (rotateItem.isKit() && showAllParts) {
                obsList = rotateItem.getKitNode().getNoneStPurchaseItemList();
            } else {
                obsList = rotateItem.getNoneStockPurchaseItemObsList();
            }

            noneStPurchaseTableView.setItems(obsList);
            if (obsList.size() > 0) {
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
