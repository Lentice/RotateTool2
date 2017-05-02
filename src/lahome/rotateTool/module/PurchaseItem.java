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

    private StringProperty kitName;
    private StringProperty partNumber;
    private StringProperty po;
    private StringProperty grDate;
    private IntegerProperty grQty;
    private IntegerProperty myQty;
    private IntegerProperty applySet;
    private StringProperty remark;

    private int rowNum;
    private StockItem stockItem;
    private RotateItem rotateItem;
    private boolean isNoneStock;

    public PurchaseItem(int rowNum, String kitName, String partNum, String po, String grDate, int grQty,
                        int myQty, int applySet, String remark) {

        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.po = new SimpleStringProperty(po);
        this.grDate = new SimpleStringProperty(grDate);
        this.grQty = new SimpleIntegerProperty(grQty);
        this.myQty = new SimpleIntegerProperty(myQty);
        this.applySet = new SimpleIntegerProperty(applySet);
        this.remark = new SimpleStringProperty(remark);

        this.rowNum = rowNum;
        this.isNoneStock = false;
    }

    public void setStock(StockItem stockItem) {
        this.stockItem = stockItem;

        stockItem.addPurchaseGrQtyTotal(getGrQty());
        stockItem.addPurchaseApQtyTotal(getMyQty());
        stockItem.addPurchaseApSetTotal(getApplySet());

        this.grQty.addListener((observable, oldValue, newValue) ->
                stockItem.addPurchaseGrQtyTotal(newValue.intValue() - oldValue.intValue()));

        this.myQty.addListener((observable, oldValue, newValue) ->
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
            rotateItem.addNoneStPurchaseApQtyTotal(getMyQty());
            rotateItem.addNoneStPurchaseApSetTotal(getApplySet());

            this.myQty.addListener((observable, oldValue, newValue) ->
                    rotateItem.addNoneStPurchaseApQtyTotal(newValue.intValue() - oldValue.intValue()));

            this.applySet.addListener((observable, oldValue, newValue) ->
                    rotateItem.addNoneStPurchaseApSetTotal(newValue.intValue() - oldValue.intValue()));
        } else {
            rotateItem.addStockGrQtyTotal(getGrQty());
            this.grQty.addListener((observable, oldValue, newValue) ->
                    rotateItem.addStockGrQtyTotal(newValue.intValue() - oldValue.intValue()));
        }
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

    public String getGrDate() {
        return grDate.get();
    }

    public StringProperty grDateProperty() {
        return grDate;
    }

    public void setGrDate(String grDate) {
        this.grDate.set(grDate);
    }

    public int getGrQty() {
        return grQty.get();
    }

    public IntegerProperty grQtyProperty() {
        return grQty;
    }

    public void setGrQty(int grQty) {
        this.grQty.set(grQty);
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

    public int getApplySet() {
        return applySet.get();
    }

    public IntegerProperty applySetProperty() {
        return applySet;
    }

    public void setApplySet(int applySet) {
        this.applySet.set(applySet);
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
}
