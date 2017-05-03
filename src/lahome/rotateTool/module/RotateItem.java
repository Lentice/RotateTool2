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

    private StringProperty kitName;
    private StringProperty partNumber;
    private IntegerProperty pmQty;
    private IntegerProperty myQty;
    private StringProperty ratio;
    private IntegerProperty applySet;
    private StringProperty remark;

    private IntegerProperty stockQtyTotal;
    private IntegerProperty stockGrQtyTotal;

    private IntegerProperty noneStPurchaseGrQtyTotal;
    private IntegerProperty noneStPurchaseApQtyTotal;
    private IntegerProperty noneStPurchaseApSetTotal;

    private IntegerProperty purchaseAllGrQtyTotal;
    private IntegerProperty purchaseAllApSetTotal;

    private boolean isKit;
    private boolean isDuplicate;
    private int rowNum;

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
        this.myQty = new SimpleIntegerProperty(0);
        this.ratio = new SimpleStringProperty(ratio);
        this.applySet = new SimpleIntegerProperty(applySet);
        this.remark = new SimpleStringProperty(remark);

        this.stockQtyTotal = new SimpleIntegerProperty(0);
        this.stockGrQtyTotal = new SimpleIntegerProperty(0);

        this.noneStPurchaseGrQtyTotal = new SimpleIntegerProperty(0);
        this.noneStPurchaseApQtyTotal = new SimpleIntegerProperty(0);
        this.noneStPurchaseApSetTotal = new SimpleIntegerProperty(0);

        this.purchaseAllGrQtyTotal = new SimpleIntegerProperty(0);
        this.purchaseAllApSetTotal = new SimpleIntegerProperty(0);

        this.isKit = !kitName.isEmpty();
        this.isDuplicate = false;
        this.rowNum = rowNum;
    }

    public boolean isKit() {
        return isKit;
    }

    public boolean isDuplicate() {
        return isDuplicate;
    }

    public void setDuplicate(List<RotateItem> duplicateRotateItems, ObservableList<StockItem> stockItemObsList) {
        this.isDuplicate = true;
        this.duplicateRotateItems = duplicateRotateItems;
        this.stockItemObsList = stockItemObsList;
        duplicateRotateItems.add(this);
    }

    public void addDuplicate(RotateItem item) {
        isDuplicate = true;
        duplicateRotateItems = new ArrayList<>();
        duplicateRotateItems.add(this);

        item.setDuplicate(duplicateRotateItems, stockItemObsList);
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

    public void addStockItem(StockItem item) {

        log.debug(String.format("Add stock: %s, %s, %s, %d, %s, %d",
                getKitName(), getPartNumber(), item.getPo(), item.getStockQty(), item.getDc(),
                item.getApplyQty()));

        addMyQty(item.getApplyQty());

        String key = item.getPo();
        item.setRotateItem(this);
        StockItem stItem = stockItems.get(key);
        if (stItem != null) {
            stItem.addDuplicate(item);
        } else {
            stockItems.put(key, item);
        }

        stockQtyTotal.set(stockQtyTotal.intValue() + item.getStockQty());
        stockItemObsList.add(item);
    }

    public ObservableList<StockItem> getStockItemObsList() {

        stockItemObsList.sort((o1, o2) -> {
            if (o2.getEarliestGrDate().isEmpty()) {
                return -1;
            } else if (o1.getEarliestGrDate().isEmpty()) {
                return 1;
            } else {
                return o1.getEarliestGrDate().compareTo(o2.getEarliestGrDate());
            }
        });
        return stockItemObsList;
    }

    public IntegerProperty myQtyProperty() {
        return myQty;
    }

    public void addMyQty(int qty) {
        myQty.set(myQty.intValue() + qty);
    }

    public StockItem getStock(String po) {
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

    public void addStockGrQtyTotal(int qty) {
        stockGrQtyTotal.set(stockGrQtyTotal.intValue() + qty);
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

    public IntegerProperty purchaseAllGrQtyTotalProperty() {
        return purchaseAllGrQtyTotal;
    }

    public void addPurchaseAllGrQtyTotal(int qty) {
        purchaseAllGrQtyTotal.set(purchaseAllGrQtyTotal.intValue() + qty);
    }

    public IntegerProperty purchaseAllApSetTotalProperty() {
        return purchaseAllApSetTotal;
    }

    public void addPurchaseAllApSetTotal(int set) {
        purchaseAllApSetTotal.set(purchaseAllApSetTotal.intValue() + set);
    }

    public void addNoneStockPurchase(PurchaseItem purchaseItem) {
        noneStPurchaseGrQtyTotal.set(noneStPurchaseGrQtyTotal.get() + purchaseItem.getGrQty());
        noneStPurchaseApQtyTotal.set(noneStPurchaseApQtyTotal.get() + purchaseItem.getApplyQty());
        noneStPurchaseApSetTotal.set(noneStPurchaseApSetTotal.get() + purchaseItem.getApplySet());

        noneStockPurchaseItemObsList.add(purchaseItem);
        purchaseItem.setRotate(this, true);
    }

    public ObservableList<PurchaseItem> getNoneStockPurchaseItemObsList() {
        noneStockPurchaseItemObsList.sort(Comparator.comparing(PurchaseItem::getGrDate));
        return noneStockPurchaseItemObsList;
    }

}
