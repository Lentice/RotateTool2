package lahome.rotateTool.view;

import com.jfoenix.controls.*;
import com.jfoenix.controls.cells.editors.IntegerTextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.TextFieldEditorBuilder;
import com.jfoenix.controls.cells.editors.base.GenericEditableTreeTableCell;
import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.collections.ObservableList;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.KeyCode;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import lahome.rotateTool.Main;
import lahome.rotateTool.Util.ExcelParser;
import lahome.rotateTool.Util.ExcelCommon;
import lahome.rotateTool.Util.ExcelSaver;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.util.prefs.Preferences;

public class RootController {
    private static final Logger log = LogManager.getLogger(RootController.class.getName());

    @FXML
    private JFXTabPane mainTab;

    @FXML
    private JFXButton saveFilesButton;

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
    private JFXTreeTableView<RotateItem> rotateTreeTableView;

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
    private JFXTreeTableView<StockItem> stockTreeTableView;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockNoColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockPartNumColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockPoColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockLotColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockDcColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockGrDateColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockStQtyColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, Number> stockApQtyColumn;

    @FXML
    private JFXTreeTableColumn<StockItem, String> stockRemarkColumn;

    @FXML
    private Label stockPmQty;

    @FXML
    private Label stockStQtyTotal;

//    @FXML
//    private Label stockGrQtyTotal;

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
    private JFXTreeTableView<PurchaseItem> purchaseTreeTableView;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchaseNoColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchasePartNumColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchasePoColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> purchaseGrDateColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> purchaseGrQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> purchaseApQtyColumn;

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
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchaseNoColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchasePartNumColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchasePoColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, String> noneStPurchaseGrDateColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> noneStPurchaseGrQtyColumn;

    @FXML
    private JFXTreeTableColumn<PurchaseItem, Number> noneStPurchaseApQtyColumn;

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
                int row = rotateTreeTableView.getSelectionModel().getSelectedIndex();
                row++;
                if (row < rotateTreeTableView.getCurrentItemsCount()) {
                    rotateTreeTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTreeTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTreeTableView.scrollToColumn(rotateApplySetColumn);
                }
                focusNode.requestFocus();

                //Stop letting it do anything else
                keyEvent.consume();
            }

            if (keyEvent.getCode() == KeyCode.F11) {

                int row = rotateTreeTableView.getSelectionModel().getSelectedIndex();
                row--;
                if (row >= 0) {
                    rotateTreeTableView.getSelectionModel().select(row, rotateApplySetColumn);
                    rotateTreeTableView.scrollTo(Math.max(row - 3, 0));
                    rotateTreeTableView.scrollToColumn(rotateApplySetColumn);
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

        ExcelCommon.setRotateConfig(
                pathAgingReport.getText(),
                Integer.valueOf(settingRotateFirstRow.getText()),
                settingRotateKitNameCol.getText(),
                settingRotatePartNumCol.getText(),
                settingRotatePmQtyCol.getText(),
                settingRotateApQtyCol.getText(),
                settingRotateRatioCol.getText(),
                settingRotateApSetCol.getText(),
                settingRotateRemarkCol.getText());

        ExcelCommon.setStockConfig(
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

        ExcelCommon.setPurchaseConfig(
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

        excelParser.loadRotateExcel();
        log.info("Process rotate table: Done");

        excelParser.loadStockExcel();
        log.info("Process stock table: Done");

        excelParser.loadPurchaseExcel();
        log.info("Process purchase table: Done");

        final TreeItem<RotateItem> rotateRoot = new RecursiveTreeItem<>(
                collection.getRotateObsList(), RecursiveTreeObject::getChildren);
        rotateTreeTableView.setRoot(rotateRoot);
        rotateTreeTableView.requestFocus();
        if (rotateTreeTableView.getCurrentItemsCount() > 0) {
            rotateTreeTableView.getSelectionModel().select(0, rotatePmQtyColumn);
        }
        rotatePartCount.setText(String.valueOf(collection.getRotateObsList().size()));

        mainTab.getSelectionModel().select(editTab);
        saveFilesButton.setVisible(true);
    }

    @FXML
    void handleSaveFiles() {
        log.info("Save File clicked");
        //excelParser.saveRotateExcel();
        //log.info("Save rotate table: Done");
//        excelParser.savePurchaseExcel();
//        log.info("Save purchase table: Done");
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
        saveFilesButton.setVisible(false);
        initialRotateTable();
        initialStockTable();
        initialPurchaseTable();
        initialNoneStockPurchaseTable();
    }

    private void initialRotateTable() {
        rotateTreeTableView.getSelectionModel().setCellSelectionEnabled(true);

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
            rotateTreeTableView.requestFocus();
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
            rotateTreeTableView.requestFocus();
        });

        rotateFilterField.textProperty().addListener((o, oldVal, newVal) -> rotateTreeTableView.setPredicate(itemProp -> {
            final RotateItem item = itemProp.getValue();
            return item.kitNameProperty().get().contains(newVal) || item.partNumberProperty().get().contains(newVal);
        }));

        rotateTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedRotateItem(newValue));

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

        stockPartNumCombo.autosize();
        stockPartNumCombo.setOnAction(e -> stockPartNumComboChanged());
    }

    private void initialStockTable() {
        stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
        stockTreeTableView.getSelectionModel().setCellSelectionEnabled(true);

        stockNoColumn.setStyle("-fx-alignment: top-left; ");
        stockPartNumColumn.setStyle("-fx-alignment: top-left; ");
        stockPoColumn.setStyle("-fx-alignment: top-left; ");
        stockLotColumn.setStyle("-fx-alignment: top-left; ");
        stockRemarkColumn.setStyle("-fx-alignment: top-left; ");

        stockNoColumn.setCellValueFactory(cellData -> {
            if (stockNoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().serialNoProperty();
            } else {
                return stockNoColumn.getComputedValue(cellData);
            }
        });
        stockNoColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        stockPartNumColumn.setCellValueFactory(cellData -> {
            if (stockPartNumColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().partNumberProperty();
            } else {
                return stockPartNumColumn.getComputedValue(cellData);
            }
        });
        stockPartNumColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        stockPoColumn.setCellValueFactory(cellData -> {
            if (stockPoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return stockPoColumn.getComputedValue(cellData);
            }
        });
        stockPoColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        stockLotColumn.setCellValueFactory(cellData -> {
            if (stockLotColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().lotProperty();
            } else {
                return stockLotColumn.getComputedValue(cellData);
            }
        });
        stockLotColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        stockDcColumn.setCellValueFactory(cellData -> {
            if (stockDcColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().dcProperty();
            } else {
                return stockDcColumn.getComputedValue(cellData);
            }
        });
        stockDcColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });


        stockGrDateColumn.setCellValueFactory(cellData -> {
            if (stockGrDateColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().earliestGrDateProperty();
            } else {
                return stockGrDateColumn.getComputedValue(cellData);
            }
        });
        stockGrDateColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        stockStQtyColumn.setCellValueFactory(cellData -> {
            if (stockStQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().stockQtyProperty();
            } else {
                return stockStQtyColumn.getComputedValue(cellData);
            }
        });
        stockStQtyColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<StockItem, Number>(new IntegerTextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(Number newValue) {
                    }
                });

        stockApQtyColumn.setCellValueFactory(cellData -> {
            if (stockApQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applyQtyProperty();
            } else {
                return stockApQtyColumn.getComputedValue(cellData);
            }
        });
        stockApQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        stockApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applyQtyProperty().set(cellData.getNewValue().intValue());
            updateStockTableTotal();
            stockTreeTableView.requestFocus();
        });

        stockRemarkColumn.setCellValueFactory(cellData -> {
            if (stockRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return stockRemarkColumn.getComputedValue(cellData);
            }
        });
        stockRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        stockRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            stockTreeTableView.requestFocus();
        });

        stockTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedStockItem(newValue));

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

        purchaseNoColumn.setStyle("-fx-alignment: top-left; ");
        purchaseNoColumn.setCellValueFactory(cellData -> {
            if (purchaseNoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().serialNoProperty();
            } else {
                return purchaseNoColumn.getComputedValue(cellData);
            }
        });

        purchasePartNumColumn.setStyle("-fx-alignment: top-left; ");
        purchasePartNumColumn.setCellValueFactory(cellData -> {
            if (purchasePartNumColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().partNumberProperty();
            } else {
                return purchasePartNumColumn.getComputedValue(cellData);
            }
        });
        purchasePartNumColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        purchasePoColumn.setStyle("-fx-alignment: top-left; ");
        purchasePoColumn.setCellValueFactory(cellData -> {
            if (purchasePoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return purchasePoColumn.getComputedValue(cellData);
            }
        });
        purchasePoColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        purchaseGrDateColumn.setCellValueFactory(cellData -> {
            if (purchaseGrDateColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grDateProperty();
            } else {
                return purchaseGrDateColumn.getComputedValue(cellData);
            }
        });
        purchaseGrDateColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        purchaseGrQtyColumn.setCellValueFactory(cellData -> {
            if (purchaseGrQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grQtyProperty();
            } else {
                return purchaseGrQtyColumn.getComputedValue(cellData);
            }
        });
        purchaseGrQtyColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, Number>(new IntegerTextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(Number newValue) {
                    }
                });

        purchaseApQtyColumn.setCellValueFactory(cellData -> {
            if (purchaseApQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applyQtyProperty();
            } else {
                return purchaseApQtyColumn.getComputedValue(cellData);
            }
        });
        purchaseApQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        purchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applyQtyProperty().set(cellData.getNewValue().intValue());
            updatePurchaseTableTotal();
            purchaseTreeTableView.requestFocus();
        });

        purchaseApplySetColumn.setCellValueFactory(cellData -> {
            if (purchaseApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return purchaseApplySetColumn.getComputedValue(cellData);
            }
        });
        purchaseApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        purchaseApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            updatePurchaseTableTotal();
            updateRotateTableTotal();
            purchaseTreeTableView.requestFocus();
        });

        purchaseRemarkColumn.setCellValueFactory(cellData -> {
            if (purchaseRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return purchaseRemarkColumn.getComputedValue(cellData);
            }
        });
        purchaseRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        purchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            purchaseTreeTableView.requestFocus();
        });

        purchaseTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedPurchaseItem(newValue));

        purchaseShowAllToggle.setOnAction(a -> {
            purchaseShowAllChanged();
        });

//        purchaseTreeTableView.setRowFactory(tv -> new JFXTreeTableRow<PurchaseItem>() {
//            @Override
//            public void updateItem(PurchaseItem item, boolean empty) {
//                super.updateItem(item, empty);
//                if (item != null && item.getPo().compareTo(purchaseSelectedPo) == 0) {
//                    setStyle("-fx-background-color: #FFF176;");
//                } else {
//                    setStyle("");
//                }
//            }
//        });

    }

    private void initialNoneStockPurchaseTable() {
        noneStPurchaseTreeTableView.setPlaceholder(new Label("ZMM 找不到符合項目"));
        noneStPurchaseTreeTableView.getSelectionModel().setCellSelectionEnabled(true);
        //noneStPurchaseTreeTableView.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);

        noneStPurchaseNoColumn.setStyle("-fx-alignment: top-left; ");
        noneStPurchaseNoColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseNoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().serialNoProperty();
            } else {
                return noneStPurchaseNoColumn.getComputedValue(cellData);
            }
        });

        noneStPurchasePartNumColumn.setStyle("-fx-alignment: top-left; ");
        noneStPurchasePartNumColumn.setCellValueFactory(cellData -> {
            if (noneStPurchasePartNumColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().getRotateItem().partNumberProperty();
            } else {
                return noneStPurchasePartNumColumn.getComputedValue(cellData);
            }
        });
        noneStPurchasePartNumColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        noneStPurchasePoColumn.setStyle("-fx-alignment: top-left; ");
        noneStPurchasePoColumn.setCellValueFactory(cellData -> {
            if (noneStPurchasePoColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().poProperty();
            } else {
                return noneStPurchasePoColumn.getComputedValue(cellData);
            }
        });
        noneStPurchasePoColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        noneStPurchaseGrDateColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseGrDateColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grDateProperty();
            } else {
                return noneStPurchaseGrDateColumn.getComputedValue(cellData);
            }
        });
        noneStPurchaseGrDateColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, String>(new TextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(String newValue) {
                    }
                });

        noneStPurchaseGrQtyColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseGrQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().grQtyProperty();
            } else {
                return noneStPurchaseGrQtyColumn.getComputedValue(cellData);
            }
        });
        noneStPurchaseGrQtyColumn.setCellFactory(cellData ->
                new GenericEditableTreeTableCell<PurchaseItem, Number>(new IntegerTextFieldEditorBuilder()) {
                    @Override
                    public void commitEdit(Number newValue) {
                    }
                });

        noneStPurchaseApQtyColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseApQtyColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applyQtyProperty();
            } else {
                return noneStPurchaseApQtyColumn.getComputedValue(cellData);
            }
        });
        noneStPurchaseApQtyColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        noneStPurchaseApQtyColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applyQtyProperty().set(cellData.getNewValue().intValue());
            updateNoneStPurchaseTableTotal();
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseApplySetColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseApplySetColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().applySetProperty();
            } else {
                return noneStPurchaseApplySetColumn.getComputedValue(cellData);
            }
        });
        noneStPurchaseApplySetColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new IntegerTextFieldEditorBuilder()));
        noneStPurchaseApplySetColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().applySetProperty().set(cellData.getNewValue().intValue());
            updateNoneStPurchaseTableTotal();
            updateRotateTableTotal();
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseRemarkColumn.setCellValueFactory(cellData -> {
            if (noneStPurchaseRemarkColumn.validateValue(cellData)) {
                return cellData.getValue().getValue().remarkProperty();
            } else {
                return noneStPurchaseRemarkColumn.getComputedValue(cellData);
            }
        });
        noneStPurchaseRemarkColumn.setCellFactory(cellData -> new GenericEditableTreeTableCell<>(new TextFieldEditorBuilder()));
        noneStPurchaseRemarkColumn.setOnEditCommit(cellData -> {
            cellData.getTreeTableView().getTreeItem(cellData.getTreeTablePosition().getRow())
                    .getValue().remarkProperty().set(cellData.getNewValue());
            noneStPurchaseTreeTableView.requestFocus();
        });

        noneStPurchaseTreeTableView.getSelectionModel().selectedItemProperty().addListener(
                (observable, oldValue, newValue) -> selectedNoneStockPurchaseItem(newValue)
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

    private void selectedRotateItem(TreeItem<RotateItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            stockTreeTableView.setRoot(null);
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

    private void selectedStockItem(TreeItem<StockItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            purchaseTreeTableView.setRoot(null);
            return;
        }

        stockPoColumn.setVisible(false);
        stockPoColumn.setVisible(true);

        StockItem stockItem = newValue.getValue();
        currentStockItem = stockItem;
        stockSelectedPo = stockItem.getPo();
        stockSelectedPart = stockItem.getRotateItem().getPartNumber();

        updatePurchaseTable(stockItem);
    }

    private void selectedPurchaseItem(TreeItem<PurchaseItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            return;
        }

        PurchaseItem purchaseItem = newValue.getValue();
        purchaseSelectedPo = purchaseItem.getPo();
    }

    private void selectedNoneStockPurchaseItem(TreeItem<PurchaseItem> newValue) {
        if (newValue == null || !newValue.isLeaf()) {
            return;
        }

        PurchaseItem purchaseItem = newValue.getValue();
        noneStPurchaseSelectedPo = purchaseItem.getPo();
    }

    private void stockPartNumComboChanged() {
        if (currentRotateItem == null)
            return;

        if (stockPartNumCombo.getItems().size() == 0)
            return;

        String name = stockPartNumCombo.getSelectionModel().getSelectedItem().trim();
        if (name.trim().isEmpty())
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
        if (currentStockItem != null)
            updatePurchaseTable(currentStockItem);
    }

    private void updateStockTable(RotateItem rotateItem) {
        TreeItem<StockItem> root = null;

        if (rotateItem.isDuplicate()) {
            root = null;
            stockTreeTableView.setPlaceholder(new Label("重複的項目!! 我無法幫妳了 orz...."));
        } else {
            ObservableList<StockItem> obsList;
            if (rotateItem.isKit() && showAllParts) {
                obsList = rotateItem.getKitNode().getStockItemList();
            } else {
                obsList = rotateItem.getStockItemObsList();
            }

            if (obsList.size() == 0) {
                stockTreeTableView.setPlaceholder(new Label("庫存找不到符合項目"));
            } else {
                root = new RecursiveTreeItem<>(obsList, RecursiveTreeObject::getChildren);
            }
        }

        stockTreeTableView.setRoot(root);
        if (stockTreeTableView.getCurrentItemsCount() > 0) {
            stockTreeTableView.getSelectionModel().select(0, stockApQtyColumn);
        }

        updateStockTableTotal();

        if (rotateItem.isKit() && showAllParts) {
            stockPmQty.setText(String.valueOf(rotateItem.getKitNode().pmQtyProperty()));
        } else {
            stockPmQty.setText(String.valueOf(rotateItem.getPmQty()));
        }

    }

    private void updatePurchaseTable(StockItem stockItem) {

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

        TreeItem<PurchaseItem> root = null;
        if (obsList.size() > 0) {
            root = new RecursiveTreeItem<>(obsList, RecursiveTreeObject::getChildren);
        }

        purchaseTreeTableView.setRoot(root);
        if (purchaseTreeTableView.getCurrentItemsCount() > 0) {
            purchaseTreeTableView.getSelectionModel().select(0, purchaseApQtyColumn);
        }
        updatePurchaseTableTotal();

        if (purchaseShowAllToggle.isSelected()) {
            purchaseTableTitle.setText("All PO in ZMM");
        } else {
            purchaseTableTitle.setText("PO:" + stockItem.getPo() + " in ZMM");
        }

    }

    private void updateNoneStPurchaseTable(RotateItem rotateItem) {
        TreeItem<PurchaseItem> noneStockRoot = null;
        if (rotateItem.isDuplicate()) {
            noneStockRoot = null;
        } else {
            ObservableList<PurchaseItem> noneStockObsList;
            if (rotateItem.isKit() && showAllParts) {
                noneStockObsList = rotateItem.getKitNode().getNoneStPurchaseItemList();
            } else {
                noneStockObsList = rotateItem.getNoneStockPurchaseItemObsList();
            }

            if (noneStockObsList.size() > 0) {
                noneStockRoot = new RecursiveTreeItem<>(noneStockObsList, RecursiveTreeObject::getChildren);
            }
        }

        noneStPurchaseTreeTableView.setRoot(noneStockRoot);
        if (noneStPurchaseTreeTableView.getCurrentItemsCount() > 0) {
            noneStPurchaseTreeTableView.getSelectionModel().select(0, noneStPurchaseApQtyColumn);
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

        TreeItem<StockItem> root = stockTreeTableView.getRoot();
        if (root != null) {
            for (TreeItem<StockItem> treeItem : root.getChildren()) {
                if (treeItem == null || treeItem.getValue() == null)
                    continue;

                StockItem stockItem = treeItem.getValue();
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
        TreeItem<PurchaseItem> root = purchaseTreeTableView.getRoot();
        if (root != null) {
            for (TreeItem<PurchaseItem> treeItem : root.getChildren()) {
                if (treeItem == null || treeItem.getValue() == null)
                    continue;

                PurchaseItem purchaseItem = treeItem.getValue();
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

        TreeItem<PurchaseItem> root = noneStPurchaseTreeTableView.getRoot();
        if (root != null) {
            for (TreeItem<PurchaseItem> treeItem : root.getChildren()) {
                if (treeItem == null || treeItem.getValue() == null)
                    continue;

                PurchaseItem purchaseItem = treeItem.getValue();
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
