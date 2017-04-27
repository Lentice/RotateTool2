package lahome.rotateTool.module;

import com.jfoenix.controls.datamodels.treetable.RecursiveTreeObject;
import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;

public class PurchaseItem extends RecursiveTreeObject<PurchaseItem> {
    private static final Logger log = LogManager.getLogger(PurchaseItem.class.getName());

    private StringProperty kitName;
    private StringProperty partNumber;
    private StringProperty po;
    private StringProperty grDate;
    private IntegerProperty grQty;

    private int rowNum;
    private StockItem stockItem;

    public PurchaseItem(int rowNum, String kitName, String partNum, String po, String grDate, int grQty) {

        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.po = new SimpleStringProperty(po);
        this.grDate = new SimpleStringProperty(grDate);
        this.grQty = new SimpleIntegerProperty(grQty);

        this.rowNum = rowNum;
    }

    public void setStockItem(StockItem item) {
        this.stockItem = item;
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
}
