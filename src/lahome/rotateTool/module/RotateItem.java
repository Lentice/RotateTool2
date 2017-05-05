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
    private boolean isDuplicate;
    private int rowNum;

    private StringProperty kitName;
    private StringProperty partNumber;
    private IntegerProperty pmQty;
    private IntegerProperty stockApplyQtyTotal;
    private StringProperty ratio;
    private IntegerProperty applySet;
    private StringProperty remark;

    private IntegerProperty stockQtyTotal;

    private IntegerProperty noneStPurchaseGrQtyTotal;
    private IntegerProperty noneStPurchaseApQtyTotal;
    private IntegerProperty noneStPurchaseApSetTotal;

    private IntegerProperty purchaseAllApSetTotal;

    private StringProperty serialNo;

    private List<RotateItem> duplicateRotateItems;

    private HashMap<String, StockItem> stockItems = new HashMap<>();
    private ObservableList<StockItem> stockItemObsList = FXCollections.observableArrayList();
    private ObservableList<PurchaseItem> noneStockPurchaseItemObsList = FXCollections.observableArrayList();


    public RotateItem(int rowNum, String kitName, String partNum, int pmQty, String ratio, int applySet, String remark) {
        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.pmQty = new SimpleIntegerProperty(pmQty);
        this.stockApplyQtyTotal = new SimpleIntegerProperty(0);
        this.ratio = new SimpleStringProperty(ratio);
        this.applySet = new SimpleIntegerProperty(applySet);
        this.remark = new SimpleStringProperty(remark);

        this.stockQtyTotal = new SimpleIntegerProperty(0);

        this.noneStPurchaseGrQtyTotal = new SimpleIntegerProperty(0);
        this.noneStPurchaseApQtyTotal = new SimpleIntegerProperty(0);
        this.noneStPurchaseApSetTotal = new SimpleIntegerProperty(0);

        this.purchaseAllApSetTotal = new SimpleIntegerProperty(0);

        this.serialNo = new SimpleStringProperty("A");

        this.isKit = !kitName.isEmpty();
        this.isDuplicate = false;
        this.rowNum = rowNum;
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

        addStockQtyTotal(item.getStockQty());
        stockItemObsList.add(item);
    }

    public void addNoneStockPurchase(PurchaseItem purchaseItem) {
        noneStPurchaseGrQtyTotal.set(noneStPurchaseGrQtyTotal.get() + purchaseItem.getGrQty());
        noneStPurchaseApQtyTotal.set(noneStPurchaseApQtyTotal.get() + purchaseItem.getApplyQty());
        noneStPurchaseApSetTotal.set(noneStPurchaseApSetTotal.get() + purchaseItem.getApplySet());

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

    public StringProperty ratioProperty() {
        return ratio;
    }


    public ObservableList<StockItem> getStockItemObsList() {
        // sort by Gr Date by default
        stockItemObsList.sort((o1, o2) -> {
            if (o2.getEarliestGrDate().isEmpty()) {
                return -1;
            } else if (o1.getEarliestGrDate().isEmpty()) {
                return 1;
            } else {
                int value = o1.getEarliestGrDate().compareTo(o2.getEarliestGrDate());
                if (value != 0)
                    return value;

                return o1.getPo().compareTo(o2.getPo());
            }
        });
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

    public IntegerProperty applySetProperty() {
        return applySet;
    }

    public StringProperty remarkProperty() {
        return remark;
    }

    public IntegerProperty stockQtyTotalProperty() {
        return stockQtyTotal;
    }

    public void addStockQtyTotal(int qty) {
        stockQtyTotal.set(stockQtyTotal.intValue() + qty);
        if (isKit) {
            assert kitNode != null;
            kitNode.addStockQtyTotal(qty);
        }
    }

    public IntegerProperty noneStPurchaseGrQtyTotalProperty() {
        return noneStPurchaseGrQtyTotal;
    }

    public IntegerProperty noneStPurchaseApQtyTotalProperty() {
        return noneStPurchaseApQtyTotal;
    }

    public void addNoneStPurchaseApQtyTotal(int qty) {
        noneStPurchaseApQtyTotal.set(noneStPurchaseApQtyTotal.intValue() + qty);
    }

    public IntegerProperty noneStPurchaseApSetTotalProperty() {
        return noneStPurchaseApSetTotal;
    }

    public void addNoneStPurchaseApSetTotal(int qty) {
        noneStPurchaseApSetTotal.set(noneStPurchaseApSetTotal.intValue() + qty);
    }

    public IntegerProperty purchaseAllApSetTotalProperty() {
        return purchaseAllApSetTotal;
    }

    public void addPurchaseAllApSetTotal(int set) {
        purchaseAllApSetTotal.set(purchaseAllApSetTotal.intValue() + set);
        if (isKit) {
            assert kitNode != null;
            kitNode.addPurchaseAllApSetTotal(set);
        }
    }

    public ObservableList<PurchaseItem> getNoneStockPurchaseItemObsList() {
        noneStockPurchaseItemObsList.sort((o1, o2) -> {
            int value = o1.getGrDate().compareTo(o2.getGrDate());
            if (value != 0)
                return value;
            return o1.getPo().compareTo(o2.getPo());
        });
        return noneStockPurchaseItemObsList;
    }

    public KitNode getKitNode() {
        return kitNode;
    }

    public void setKitNode(KitNode kitNode) {
        this.kitNode = kitNode;
    }

    public void setSerialNo(String serialNo) {
        this.serialNo.set(serialNo);
    }

    public StringProperty serialNoProperty() {
        return serialNo;
    }
}
