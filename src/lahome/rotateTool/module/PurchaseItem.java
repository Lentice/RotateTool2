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
    private StockItem stockItem;
    private RotateItem rotateItem;
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

        stockItem.addPurchaseGrQtyTotal(getGrQty());
        stockItem.addPurchaseApQtyTotal(getApplyQty());
        stockItem.addPurchaseApSetTotal(getApplySet());

        this.grQty.addListener((observable, oldValue, newValue) ->
                stockItem.addPurchaseGrQtyTotal(newValue.intValue() - oldValue.intValue()));

        this.applyQty.addListener((observable, oldValue, newValue) ->
                stockItem.addPurchaseApQtyTotal(newValue.intValue() - oldValue.intValue()));

        this.applySet.addListener((observable, oldValue, newValue) ->
                stockItem.addPurchaseApSetTotal(newValue.intValue() - oldValue.intValue()));
    }

    public void setRotate(RotateItem rotateItem, boolean isNoneStock) {
        this.isNoneStock = isNoneStock;
        this.rotateItem = rotateItem;

        rotateItem.addPurchaseAllGrQtyTotal(getGrQty());
        rotateItem.addPurchaseAllApSetTotal(getApplySet());
        this.applySet.addListener((observable, oldValue, newValue) ->
                rotateItem.addPurchaseAllApSetTotal(newValue.intValue() - oldValue.intValue()));
        this.grQty.addListener((observable, oldValue, newValue) ->
                rotateItem.addPurchaseAllGrQtyTotal(newValue.intValue() - oldValue.intValue()));

        if (isNoneStock) {
            rotateItem.addNoneStPurchaseApQtyTotal(getApplyQty());
            rotateItem.addNoneStPurchaseApSetTotal(getApplySet());

            this.applyQty.addListener((observable, oldValue, newValue) ->
                    rotateItem.addNoneStPurchaseApQtyTotal(newValue.intValue() - oldValue.intValue()));

            this.applySet.addListener((observable, oldValue, newValue) ->
                    rotateItem.addNoneStPurchaseApSetTotal(newValue.intValue() - oldValue.intValue()));
        } else {
            rotateItem.addStockGrQtyTotal(getGrQty());
            this.grQty.addListener((observable, oldValue, newValue) ->
                    rotateItem.addStockGrQtyTotal(newValue.intValue() - oldValue.intValue()));
        }
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

    public StringProperty remarkProperty() {
        return remark;
    }

}
