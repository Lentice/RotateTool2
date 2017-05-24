package lahome.rotateTool.module;

import gnu.trove.map.hash.THashMap;
import javafx.beans.property.*;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

@SuppressWarnings("WeakerAccess")
public class RotateItem {
    private static final Logger log = LogManager.getLogger(RotateItem.class.getName());

    private boolean isKit;
    private KitNode kitNode;
    private boolean isFirstPartOfKit = false;
    private boolean isDuplicate = false;
    private boolean isRotateValid = true;
    private int rowNum;

    private ReadOnlyStringWrapper kitName;
    private ReadOnlyStringWrapper partNumber;
    private ReadOnlyIntegerWrapper backlog;
    private ReadOnlyIntegerWrapper pmQty;
    private ReadOnlyIntegerWrapper stockApplyQtyTotal;
    private ReadOnlyStringWrapper ratio;
    private ReadOnlyIntegerWrapper applySet;
    private StringProperty remark;

    private ReadOnlyStringWrapper serialNo;
    private int purchasesApplyQtyTotal;

    private List<RotateItem> duplicateRotateItems;

    private THashMap<String, StockItem> stockItemMap = new THashMap<>();
    private ObservableList<StockItem> stockItemObsList = FXCollections.observableArrayList();
    private ObservableList<PurchaseItem> noneStockPurchaseItemObsList = FXCollections.observableArrayList();

    public RotateItem(int rowNum, String kitName, String partNum, int backlog, int pmQty, String ratio, String remark) {
        if (kitName == null)
            kitName = "";

        this.kitName = new ReadOnlyStringWrapper(kitName);
        this.partNumber = new ReadOnlyStringWrapper(partNum);
        this.backlog = new ReadOnlyIntegerWrapper(backlog);
        this.pmQty = new ReadOnlyIntegerWrapper(pmQty);
        this.stockApplyQtyTotal = new ReadOnlyIntegerWrapper(0);
        this.ratio = new ReadOnlyStringWrapper(ratio);
        this.applySet = new ReadOnlyIntegerWrapper(0);
        this.remark = new SimpleStringProperty(remark);

        this.serialNo = new ReadOnlyStringWrapper("A");

        this.purchasesApplyQtyTotal = 0;

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

//        log.debug(String.format("Add stock: %s, %s, %s, %s, %d, %s, %d",
//                getKitName(), getPartNumber(), getBacklog(), item.getPo(), item.getStockQty(), item.getDc(),
//                item.getApplyQty()));

        item.setRotateItem(this);

        String key = item.getPo();
        StockItem stItem = stockItemMap.get(key);
        if (stItem != null) {
            stItem.addDuplicate(item);
        } else {
            stockItemMap.put(key, item);
        }

        stockItemObsList.add(item);
    }

    public void addNoneStockPurchase(PurchaseItem purchaseItem) {
        noneStockPurchaseItemObsList.add(purchaseItem);
        purchaseItem.setRotateAndStock(this, null);
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

    public ReadOnlyStringProperty kitNameProperty() {
        return kitName.getReadOnlyProperty();
    }

    public String getPartNumber() {
        return partNumber.get();
    }

    public ReadOnlyStringProperty partNumberProperty() {
        return partNumber.getReadOnlyProperty();
    }

    public int getBacklog() {
        return backlog.get();
    }

    public ReadOnlyIntegerProperty backlogProperty() {
        return backlog.getReadOnlyProperty();
    }

    public int getPmQty() {
        return pmQty.get();
    }

    public ReadOnlyIntegerProperty pmQtyProperty() {
        return pmQty.getReadOnlyProperty();
    }

    public int getRatio() {
        try {
            return Integer.valueOf(ratio.get());
        } catch (NumberFormatException e) {
            return 0;
        }
    }

    public ReadOnlyStringProperty ratioProperty() {
        return ratio.getReadOnlyProperty();
    }


    public ObservableList<StockItem> getStockItemObsList() {
        stockItemObsList.sort(StockItem::compareTo);
        return stockItemObsList;
    }

    public int getStockApplyQtyTotal() {
        return stockApplyQtyTotal.get();
    }

    public ReadOnlyIntegerProperty stockApplyQtyTotalProperty() {
        return stockApplyQtyTotal.getReadOnlyProperty();
    }

    public void addStockApplyQtyTotal(int qty) {
        stockApplyQtyTotal.set(stockApplyQtyTotal.intValue() + qty);
    }

    public int getPurchasesApplyQtyTotal() {
        return purchasesApplyQtyTotal;
    }

    public void addPurchasesApplyQtyTotal(int qty) {
        purchasesApplyQtyTotal += qty;
    }


    public StockItem getStockByPo(String po) {
        return stockItemMap.get(po);
    }

    public int getApplySet() {
        return applySet.get();
    }

    public ReadOnlyIntegerProperty applySetProperty() {
        return applySet.getReadOnlyProperty();
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

    public ReadOnlyStringProperty serialNoProperty() {
        return serialNo.getReadOnlyProperty();
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
                list.addAll(stockItem.getPurchaseItemList());
            }
        }

        list.sort(PurchaseItem::compareTo);
        return list;
    }

    public List<PurchaseItem> getStAndNoneStPurchaseItemList() {
        List<PurchaseItem> list = new ArrayList<>();

        for (StockItem stockItem : stockItemObsList) {
            if (stockItem.isMainStockItem()) {
                list.addAll(stockItem.getPurchaseItemList());
            }
        }
        list.addAll(noneStockPurchaseItemObsList);

        list.sort(PurchaseItem::compareTo);
        return list;
    }
}
