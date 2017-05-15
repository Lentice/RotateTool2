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

public class StockItem {
    private static final Logger log = LogManager.getLogger(StockItem.class.getName());

    private StringProperty po;
    private IntegerProperty stockQty;
    private StringProperty lot;
    private IntegerProperty dc;
    private StringProperty earliestGrDate;
    private IntegerProperty applyQty;
    private StringProperty remark;

    private int rowNum;
    private RotateItem rotateItem;

    private StockItem firstStockItem = this;
    private boolean isMainStockItem = true;

    private boolean isDuplicate;
    private List<StockItem> stockDuplicatePoItemList;

    private ObservableList<PurchaseItem> purchaseItemList = FXCollections.observableArrayList();

    public StockItem(int rowNum, String po,
                     int stockQty, String lot, int dc, int applyQty, String remark) {

        this.po = new SimpleStringProperty(po);
        this.stockQty = new SimpleIntegerProperty(stockQty);
        this.lot = new SimpleStringProperty(lot);
        this.dc = new SimpleIntegerProperty(dc);
        this.earliestGrDate = new SimpleStringProperty("");
        this.applyQty = new SimpleIntegerProperty(applyQty);
        this.remark = new SimpleStringProperty(remark);

        this.rowNum = rowNum;
    }

    public void setRotateItem(RotateItem rotateItem) {
        this.rotateItem = rotateItem;

        this.rotateItem.addStockApplyQtyTotal(getApplyQty());
        this.applyQty.addListener((observable, oldValue, newValue) -> {
            if (this.rotateItem != null) {
                this.rotateItem.addStockApplyQtyTotal(newValue.intValue() - oldValue.intValue());
            }
        });
    }

    public void addDuplicate(StockItem item) {

        assert isMainStockItem;
        if (!isDuplicate) {
            isDuplicate = true;
            stockDuplicatePoItemList = new ArrayList<>();
            stockDuplicatePoItemList.add(this);
        }

        stockDuplicatePoItemList.add(item);
        item.setDuplicate(this);
    }

    public void setDuplicate(StockItem firstItem) {

        this.isDuplicate = true;
        this.firstStockItem = firstItem;
        this.isMainStockItem = false;

        this.stockDuplicatePoItemList = firstItem.getStockDuplicatePoItemList();
        this.purchaseItemList = firstItem.getPurchaseItemList();
        this.earliestGrDate = firstItem.earliestGrDateProperty();
    }

    public int compareTo(StockItem item) {

        int value;
        if (item.getEarliestGrDate().isEmpty() && this.getEarliestGrDate().isEmpty()) {
            value = this.getPo().compareTo(item.getPo());
            if (value != 0)
                return value;

            value = this.getRotateItem().getPartNumber().compareTo(item.getRotateItem().getPartNumber());
            if (value != 0)
                return value;

            return this.getLot().compareTo(item.getLot());
        } else if (item.getEarliestGrDate().isEmpty()) {
            return -1;
        } else if (this.getEarliestGrDate().isEmpty()) {
            return 1;
        } else {
            value = this.getEarliestGrDate().compareTo(item.getEarliestGrDate());
            if (value != 0)
                return value;

            value = this.getPo().compareTo(item.getPo());
            if (value != 0)
                return value;

            value = this.getRotateItem().getPartNumber().compareTo(item.getRotateItem().getPartNumber());
            if (value != 0)
                return value;

            return this.getLot().compareTo(item.getLot());
        }
    }

    public boolean isMainStockItem() {
        return isMainStockItem;
    }

    public RotateItem getRotateItem() {
        return rotateItem;
    }

    public int getRowNum() {
        return rowNum;
    }

    public String getPo() {
        return po.get();
    }

    public StringProperty poProperty() {
        return po;
    }

    public int getStockQty() {
        return stockQty.get();
    }

    public IntegerProperty stockQtyProperty() {
        return stockQty;
    }

    public String getLot() {
        return lot.get();
    }

    public StringProperty lotProperty() {
        return lot;
    }

    public int getDc() {
        return dc.get();
    }

    public IntegerProperty dcProperty() {
        return dc;
    }

    public String getEarliestGrDate() {
        return earliestGrDate.get();
    }

    public StringProperty earliestGrDateProperty() {
        return earliestGrDate;
    }

    public int getApplyQty() {
        return applyQty.get();
    }

    public IntegerProperty applyQtyProperty() {
        return applyQty;
    }

    public String getRemark() {
        return remark.get();
    }

    public StringProperty remarkProperty() {
        return remark;
    }

    public void addPurchaseItem(PurchaseItem item) {
        log.debug(String.format("Add purchase: %s, %s, %s, %s, %d",
                rotateItem.getKitName(), rotateItem.getPartNumber(), item.getPo(), item.getGrDate(), item.getGrQty()));

        Date grDate = DateUtil.parse(item.getGrDate());
        Date earliestDate = DateUtil.parse(getEarliestGrDate());
        if (grDate != null && (earliestDate == null || grDate.before(earliestDate))) {
            this.earliestGrDate.set(item.getGrDate());
        }

        purchaseItemList.add(item);
        item.setStock(this);
        item.setRotate(rotateItem, false);
    }

    public List<StockItem> getStockDuplicatePoItemList() {
        return stockDuplicatePoItemList;
    }

    public ObservableList<PurchaseItem> getPurchaseItemList() {
        purchaseItemList.sort(PurchaseItem::compareTo);
        return purchaseItemList;
    }

    public String getPartNumber() {
        return rotateItem.getPartNumber();
    }
}
