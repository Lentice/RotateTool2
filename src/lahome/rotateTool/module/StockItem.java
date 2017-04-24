package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
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
    private IntegerProperty dc;
    private IntegerProperty myQty;

    private Row row;
    private RotateItem rotateItem;

    private List<PurchaseItem> purchaseItems = new ArrayList<>();

    public StockItem(Row row, String kitName, String partNum, String po,
                     int stockQty, int dc, int myQty) {

        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.po = new SimpleStringProperty(po);
        this.stockQty = new SimpleIntegerProperty(stockQty);
        this.dc = new SimpleIntegerProperty(dc);
        this.myQty = new SimpleIntegerProperty(myQty);

        this.row = row;
    }

    public void setRotateItem(RotateItem rotateItem) {
        this.rotateItem = rotateItem;

        this.myQty.addListener((observable, oldValue, newValue) -> {
            if (this.rotateItem != null) {
                this.rotateItem.addMyQty(newValue.intValue() - oldValue.intValue());
            }
        });
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

    public int getDc() {
        return dc.get();
    }

    public IntegerProperty dcProperty() {
        return dc;
    }

    public void setDc(int dc) {
        this.dc.set(dc);
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

    public void addPurchaseItem(PurchaseItem item) {
        log.debug(String.format("Add purchase: %s, %s, %s, %d, %d, %d",
                item.getKitName(), item.getPartNumber(), item.getPo(), item.getGrDate(), item.getGrQty()));

        purchaseItems.add(item);
        item.setStockItem(this);
    }
}
