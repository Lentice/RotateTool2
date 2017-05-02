package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import lahome.rotateTool.Util.DateUtil;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by Administrator on 2017/4/15.
 */
public class StockItem extends RecursiveTreeObject<StockItem> {
    private static final Logger log = LogManager.getLogger(StockItem.class.getName());

    private StringProperty kitName;
    private StringProperty partNumber;
    private StringProperty po;
    private IntegerProperty stockQty;
    private StringProperty lot;
    private StringProperty dc;
    private StringProperty earliestGrDate;
    private IntegerProperty myQty;
    private StringProperty remark;

    private int rowNum;
    private RotateItem rotateItem;

    private boolean isDuplicate;
    private StockItem firstStockItem;
    private List<StockItem> duplicateStockItems;
    private IntegerProperty duplicatePoStockQtyTotal;
    private IntegerProperty duplicatePoMyQtyTotal;

    private IntegerProperty purchaseGrQtyTotal;
    private IntegerProperty purchaseApQtyTotal;
    private IntegerProperty purchaseApSetTotal;

    private ObservableList<PurchaseItem> purchaseItems = FXCollections.observableArrayList();

    public StockItem(int rowNum, String kitName, String partNum, String po,
                     int stockQty, String lot, String dc, int myQty, String remark) {

        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.po = new SimpleStringProperty(po);
        this.stockQty = new SimpleIntegerProperty(stockQty);
        this.lot = new SimpleStringProperty(lot);
        this.dc = new SimpleStringProperty(dc);
        this.earliestGrDate = new SimpleStringProperty("");
        this.myQty = new SimpleIntegerProperty(myQty);
        this.remark = new SimpleStringProperty(remark);

        this.purchaseGrQtyTotal = new SimpleIntegerProperty(0);
        this.purchaseApQtyTotal = new SimpleIntegerProperty(0);
        this.purchaseApSetTotal = new SimpleIntegerProperty(0);

        this.duplicatePoStockQtyTotal = new SimpleIntegerProperty(stockQty);
        this.duplicatePoMyQtyTotal = new SimpleIntegerProperty(myQty);

        this.rowNum = rowNum;
    }

    public void setRotateItem(RotateItem rotateItem) {
        this.rotateItem = rotateItem;

        this.myQty.addListener((observable, oldValue, newValue) -> {
            if (this.rotateItem != null) {
                this.rotateItem.addMyQty(newValue.intValue() - oldValue.intValue());
            }
        });
    }

    public void setDuplicate(StockItem firstItem) {

        this.isDuplicate = true;
        this.firstStockItem = firstItem;

        this.duplicateStockItems = firstItem.getDuplicateStockItems();
        this.purchaseItems = firstItem.getPurchaseItems();
        this.earliestGrDate = firstItem.earliestGrDateProperty();
        this.purchaseGrQtyTotal = firstItem.purchaseGrQtyTotalProperty();
        this.purchaseApQtyTotal = firstItem.purchaseApQtyTotalProperty();
        this.purchaseApSetTotal = firstItem.purchaseApSetTotalProperty();
        this.duplicatePoStockQtyTotal = firstItem.duplicatePoStockQtyTotalProperty();
        this.duplicatePoMyQtyTotal = firstItem.duplicatePoMyQtyTotalProperty();
        this.duplicateStockItems.add(this);

        this.myQty.addListener((observable, oldValue, newValue) -> {
            firstItem.addDuplicatePoMyQtyTotal(newValue.intValue() - oldValue.intValue());
        });
    }

    public void addDuplicate(StockItem item) {

        if (!isDuplicate) {
            firstStockItem = this;
            this.myQty.addListener((observable, oldValue, newValue) -> {
                this.addDuplicatePoMyQtyTotal(newValue.intValue() - oldValue.intValue());
            });
        }

        isDuplicate = true;

        duplicatePoStockQtyTotal.set(duplicatePoStockQtyTotal.get() + item.getStockQty());
        addDuplicatePoMyQtyTotal(item.getMyQty());

        duplicateStockItems = new ArrayList<>();
        duplicateStockItems.add(this);

        item.setDuplicate(this);
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

    public void setKitName(String kitName) {
        this.kitName.set(kitName);
    }

    public String getPartNumber() {
        return partNumber.get();
    }

    public StringProperty partNumberProperty() {
        return partNumber;
    }

    public void setPartNumber(String partNumber) {
        this.partNumber.set(partNumber);
    }

    public String getPo() {
        return po.get();
    }

    public StringProperty poProperty() {
        return po;
    }

    public void setPo(String po) {
        this.po.set(po);
    }

    public int getStockQty() {
        return stockQty.get();
    }

    public IntegerProperty stockQtyProperty() {
        return stockQty;
    }

    public void setStockQty(int stockQty) {
        this.stockQty.set(stockQty);
    }

    public String getLot() {
        return lot.get();
    }

    public StringProperty lotProperty() {
        return lot;
    }

    public void setLot(String lot) {
        this.lot.set(lot);
    }

    public String getDc() {
        return dc.get();
    }

    public StringProperty dcProperty() {
        return dc;
    }

    public void setDc(String dc) {
        this.dc.set(dc);
    }

    public String getEarliestGrDate() {
        return earliestGrDate.get();
    }

    public StringProperty earliestGrDateProperty() {
        return earliestGrDate;
    }

    public void setEarliestGrDate(String earliestGrDate) {
        this.earliestGrDate.set(earliestGrDate);
    }

    public int getMyQty() {
        return myQty.get();
    }

    public IntegerProperty myQtyProperty() {
        return myQty;
    }

    public void setMyQty(int myQty) {
        this.myQty.set(myQty);
    }

    public int getPurchaseGrQtyTotal() {
        return purchaseGrQtyTotal.get();
    }

    public IntegerProperty purchaseGrQtyTotalProperty() {
        return purchaseGrQtyTotal;
    }

    public void setPurchaseGrQtyTotal(int purchaseGrQtyTotal) {
        this.purchaseGrQtyTotal.set(purchaseGrQtyTotal);
    }

    public int getPurchaseApQtyTotal() {
        return purchaseApQtyTotal.get();
    }

    public IntegerProperty purchaseApQtyTotalProperty() {
        return purchaseApQtyTotal;
    }

    public void setPurchaseApQtyTotal(int purchaseApQtyTotal) {
        this.purchaseApQtyTotal.set(purchaseApQtyTotal);
    }

    public void addPurchaseApQtyTotal(int qty) {
        purchaseApQtyTotal.set(purchaseApQtyTotal.intValue() + qty);
    }

    public int getPurchaseApSetTotal() {
        return purchaseApSetTotal.get();
    }

    public IntegerProperty purchaseApSetTotalProperty() {
        return purchaseApSetTotal;
    }

    public void setPurchaseApSetTotal(int purchaseApSetTotal) {
        this.purchaseApSetTotal.set(purchaseApSetTotal);
    }

    public void addPurchaseApSetTotal(int qty) {
        purchaseApSetTotal.set(purchaseApSetTotal.intValue() + qty);
    }

    public String getRemark() {
        return remark.get();
    }

    public StringProperty remarkProperty() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark.set(remark);
    }

    public int getDuplicatePoStockQtyTotal() {
        return duplicatePoStockQtyTotal.get();
    }

    public IntegerProperty duplicatePoStockQtyTotalProperty() {
        return duplicatePoStockQtyTotal;
    }

    public void setDuplicatePoStockQtyTotal(int duplicatePoStockQtyTotal) {
        this.duplicatePoStockQtyTotal.set(duplicatePoStockQtyTotal);
    }

    public int getDuplicatePoMyQtyTotal() {
        return duplicatePoMyQtyTotal.get();
    }

    public IntegerProperty duplicatePoMyQtyTotalProperty() {
        return duplicatePoMyQtyTotal;
    }

    public void addDuplicatePoMyQtyTotal(int qty) {
        duplicatePoMyQtyTotal.set(duplicatePoMyQtyTotal.intValue() + qty);
    }

    public void addPurchaseItem(PurchaseItem item) {
        log.debug(String.format("Add purchase: %s, %s, %s, %s, %d",
                item.getKitName(), item.getPartNumber(), item.getPo(), item.getGrDate(), item.getGrQty()));

        purchaseGrQtyTotal.set(purchaseGrQtyTotal.get() + item.getGrQty());
        purchaseApQtyTotal.set(purchaseApQtyTotal.get() + item.getMyQty());
        purchaseApSetTotal.set(purchaseApSetTotal.get() + item.getApplySet());


        rotateItem.setGrQtyTotal(rotateItem.getGrQtyTotal() + item.getGrQty());

        Date purchaseDate = DateUtil.parse(item.getGrDate());
        Date earliestDate = DateUtil.parse(earliestGrDate.get());
        if (purchaseDate != null && earliestDate == null) {
            earliestGrDate.set(item.getGrDate());
        } else if (purchaseDate != null && purchaseDate.before(earliestDate)) {
            earliestGrDate.set(item.getGrDate());
        }

        purchaseItems.add(item);
        item.setStockItem(this);
    }

    public List<StockItem> getDuplicateStockItems() {
        return duplicateStockItems;
    }

    public ObservableList<PurchaseItem> getPurchaseItems() {
        return purchaseItems;
    }

    public StockItem getFirstStockItem() {
        return firstStockItem;
    }
}
