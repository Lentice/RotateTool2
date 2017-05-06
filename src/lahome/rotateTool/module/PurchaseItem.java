package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class PurchaseItem extends RecursiveTreeObject<PurchaseItem> {
    private static final Logger log = LogManager.getLogger(PurchaseItem.class.getName());

    private StringProperty po;
    private StringProperty grDate;
    private IntegerProperty grQty;
    private IntegerProperty applyQty;
    private IntegerProperty applySet;
    private StringProperty remark;

    private int rowNum;
    public StockItem stockItem;
    public RotateItem rotateItem;
    private boolean isNoneStock;

    public PurchaseItem(int rowNum, String po, String grDate, int grQty,
                        int applyQty, int applySet, String remark) {

        this.po = new SimpleStringProperty(po);
        this.grDate = new SimpleStringProperty(grDate);
        this.grQty = new SimpleIntegerProperty(grQty);
        this.applyQty = new SimpleIntegerProperty(applyQty);
        this.applySet = new SimpleIntegerProperty(applySet);
        this.remark = new SimpleStringProperty(remark);

        this.rowNum = rowNum;
        this.isNoneStock = false;
    }

    public void close() {
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

    public StringProperty poProperty() {
        return po;
    }

    public String getGrDate() {
        return grDate.get();
    }

    public StringProperty grDateProperty() {
        return grDate;
    }

    public int getGrQty() {
        return grQty.get();
    }

    public IntegerProperty grQtyProperty() {
        return grQty;
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

}
