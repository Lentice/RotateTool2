package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class PurchaseItem {
    private static final Logger log = LogManager.getLogger(PurchaseItem.class.getName());

    private ReadOnlyStringWrapper po;
    private ReadOnlyStringWrapper grDate;
    private ReadOnlyIntegerWrapper grQty;
    private IntegerProperty applyQty;
    private IntegerProperty applySet;
    private StringProperty remark;

    private int rowNum;
    private StockItem stockItem;
    private RotateItem rotateItem;
    private boolean isNoneStock;

    public PurchaseItem(int rowNum, String po, String grDate, int grQty,
                        int applyQty, String remark) {

        this.po = new ReadOnlyStringWrapper(po);
        this.grDate = new ReadOnlyStringWrapper(grDate);
        this.grQty = new ReadOnlyIntegerWrapper(grQty);
        this.applyQty = new SimpleIntegerProperty(applyQty);
        this.applySet = new SimpleIntegerProperty(0);
        this.remark = new SimpleStringProperty(remark);

        this.rowNum = rowNum;
        this.isNoneStock = false;

    }

    public void clear() {
        this.po = null;
        this.grDate = null;
        this.grQty = null;
        this.applyQty = null;
        this.applySet = null;
        this.remark = null;
    }

    public void setStock(StockItem stockItem) {
        this.stockItem = stockItem;
    }

    public void setRotate(RotateItem rotateItem, boolean isNoneStock) {
        this.isNoneStock = isNoneStock;
        this.rotateItem = rotateItem;

        updateApplySet();
        this.rotateItem.addPurchasesApplyQtyTotal(getApplyQty());

        this.applyQty.addListener((observable, oldValue, newValue) -> {
            updateApplySet();

            if (this.rotateItem != null) {
                this.rotateItem.addPurchasesApplyQtyTotal(newValue.intValue() - oldValue.intValue());
            }
        });
    }

    private void updateApplySet() {
        if (rotateItem.isKit() && rotateItem.isRotateValid() && rotateItem.isFirstPartOfKit()) {
            int ratioValue = rotateItem.getRatio();
            if (ratioValue > 0) {
                applySet.set(getApplyQty() / ratioValue);
            }
        }
    }

    public int compareTo(PurchaseItem item) {
        int value = this.getGrDate().compareTo(item.getGrDate());
        if (value != 0)
            return value;

        value = this.getPo().compareTo(item.getPo());
        if (value != 0)
            return value;

        return this.getRotateItem().getPartNumber().compareTo(item.getRotateItem().getPartNumber());
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

    public ReadOnlyStringProperty poProperty() {
        return po.getReadOnlyProperty();
    }

    public String getGrDate() {
        return grDate.get();
    }

    public ReadOnlyStringProperty grDateProperty() {
        return grDate.getReadOnlyProperty();
    }

    public int getGrQty() {
        return grQty.get();
    }

    public ReadOnlyIntegerProperty grQtyProperty() {
        return grQty.getReadOnlyProperty();
    }

    public int getApplyQty() {
        return applyQty.get();
    }

    public IntegerProperty applyQtyProperty() {
        return applyQty;
    }

    public int getApplySet() {
        return applySet.get();
    }

    public IntegerProperty applySetProperty() {
        return applySet;
    }

    public String getRemark() {
        return remark.get();
    }

    public StringProperty remarkProperty() {
        return remark;
    }

    public String getPartNumber() {
        return rotateItem.getPartNumber();
    }
}
