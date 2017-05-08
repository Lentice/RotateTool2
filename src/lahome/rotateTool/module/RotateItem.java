package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.*;

public class RotateItem extends RecursiveTreeObject<RotateItem> {
    private static final Logger log = LogManager.getLogger(RotateItem.class.getName());

    private boolean isKit;
    private KitNode kitNode;
    private boolean isDuplicate = false;
    private boolean isRotateValid = true;
    private int rowNum;

    private StringProperty kitName;
    private StringProperty partNumber;
    private IntegerProperty pmQty;
    private IntegerProperty stockApplyQtyTotal;
    private StringProperty ratio;
    private IntegerProperty applySet;
    private StringProperty remark;

    private boolean isFirstPartOfKit = false;
    private StringProperty serialNo;

    private List<RotateItem> duplicateRotateItems;

    private HashMap<String, StockItem> stockItems = new HashMap<>();
    private ObservableList<StockItem> stockItemObsList = FXCollections.observableArrayList();
    private ObservableList<PurchaseItem> noneStockPurchaseItemObsList = FXCollections.observableArrayList();


    public RotateItem(int rowNum, String kitName, String partNum, int pmQty, String ratio, String remark) {
        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.pmQty = new SimpleIntegerProperty(pmQty);
        this.stockApplyQtyTotal = new SimpleIntegerProperty(0);
        this.ratio = new SimpleStringProperty(ratio);
        this.applySet = new SimpleIntegerProperty(0);
        this.remark = new SimpleStringProperty(remark);

        this.serialNo = new SimpleStringProperty("A");

        this.isKit = !kitName.isEmpty();
        this.rowNum = rowNum;
    }

    public void updateApplySet() {
        if (isKit && isRotateValid && isFirstPartOfKit) {
            int ratioValue = getRatio();
            if (ratioValue > 0) {
                applySet.set(getStockApplyQtyTotal() / ratioValue);
            }
        }
    }

    public void applySetListenerStart() {
        updateApplySet();
        if (isKit && isRotateValid && isFirstPartOfKit) {
            stockApplyQtyTotal.addListener((observable, oldValue, newValue) -> updateApplySet());
        }
    }

    public void setDuplicate(List<RotateItem> duplicateRotateItems, ObservableList<StockItem> stockItemObsList) {
        this.isDuplicate = true;
        this.duplicateRotateItems = duplicateRotateItems;
        this.stockItemObsList = stockItemObsList;
        duplicateRotateItems.add(this);
    }

    public void addDuplicate(RotateItem item) {
        if (!isDuplicate) {
            isDuplicate = true;
            duplicateRotateItems = new ArrayList<>();
            duplicateRotateItems.add(this);
        }

        item.setDuplicate(duplicateRotateItems, stockItemObsList);
    }


    public void addStockItem(StockItem item) {

        log.debug(String.format("Add stock: %s, %s, %s, %d, %s, %d",
                getKitName(), getPartNumber(), item.getPo(), item.getStockQty(), item.getDc(),
                item.getApplyQty()));

        item.setRotateItem(this);

        String key = item.getPo();
        StockItem stItem = stockItems.get(key);
        if (stItem != null) {
            stItem.addDuplicate(item);
        } else {
            stockItems.put(key, item);
        }

        stockItemObsList.add(item);
    }

    public void addNoneStockPurchase(PurchaseItem purchaseItem) {
        noneStockPurchaseItemObsList.add(purchaseItem);
        purchaseItem.setRotate(this, true);
    }

    public int getRowNum() {
        return rowNum;
    }

    public boolean isKit() {
        return isKit;
    }

    public boolean isDuplicate() {
        return isDuplicate;
    }

    public String getKitName() {
        return kitName.get();
    }

    public StringProperty kitNameProperty() {
        return kitName;
    }

    public String getPartNumber() {
        return partNumber.get();
    }

    public StringProperty partNumberProperty() {
        return partNumber;
    }

    public int getPmQty() {
        return pmQty.get();
    }

    public IntegerProperty pmQtyProperty() {
        return pmQty;
    }

    public int getRatio() {
        try {
            return Integer.valueOf(ratio.get());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    public StringProperty ratioProperty() {
        return ratio;
    }


    public ObservableList<StockItem> getStockItemObsList() {
        stockItemObsList.sort(StockItem::compareTo);
        return stockItemObsList;
    }

    public int getStockApplyQtyTotal() {
        return stockApplyQtyTotal.get();
    }

    public IntegerProperty stockApplyQtyTotalProperty() {
        return stockApplyQtyTotal;
    }

    public void addStockApplyQtyTotal(int qty) {
        stockApplyQtyTotal.set(stockApplyQtyTotal.intValue() + qty);
    }

    public StockItem getStockByPo(String po) {
        return stockItems.get(po);
    }

    public int getApplySet() {
        return applySet.get();
    }

    public IntegerProperty applySetProperty() {
        return applySet;
    }

    public StringProperty remarkProperty() {
        return remark;
    }

    public ObservableList<PurchaseItem> getNoneStockPurchaseItemObsList() {
        noneStockPurchaseItemObsList.sort(PurchaseItem::compareTo);

        return noneStockPurchaseItemObsList;
    }

    public KitNode getKitNode() {
        return kitNode;
    }

    public void setKitNode(KitNode kitNode) {
        this.kitNode = kitNode;
    }


    public String getRemark() {
        return remark.get();
    }

    public String getSerialNo() {
        return serialNo.get();
    }

    public void setSerialNo(String serialNo) {
        this.serialNo.set(serialNo);
    }
    public StringProperty serialNoProperty() {
        return serialNo;
    }

    public boolean isFirstPartOfKit() {
        return isFirstPartOfKit;
    }

    public void setFirstPartOfKit(boolean firstPartOfKit) {
        isFirstPartOfKit = firstPartOfKit;
    }

    public boolean isRotateValid() {
        return isRotateValid;
    }

    public void setRotateValid(boolean rotateValid) {
        this.isRotateValid = rotateValid;
    }

    public ObservableList<PurchaseItem> getPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();

            for (StockItem stockItem : stockItemObsList) {
                if (stockItem.isMainStockItem()) {
                    list.addAll(stockItem.getPurchaseItems());
                }
            }

        list.sort(PurchaseItem::compareTo);
        return list;
    }
}
