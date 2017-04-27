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
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class RotateItem extends RecursiveTreeObject<RotateItem> {
    private static final Logger log = LogManager.getLogger(RotateItem.class.getName());

    private StringProperty kitName;
    private StringProperty partNumber;
    private IntegerProperty pmQty;
    private IntegerProperty myQty;

    private boolean isKit;
    private boolean isDuplicate;
    private int rowNum;

    private List<RotateItem> duplicateRotateItems;

    private HashMap<String, StockItem> stockItems = new HashMap<>();

    private ObservableList<StockItem> stockItemObsList = FXCollections.observableArrayList();

    private ObservableList<PurchaseItem> noneStockPurchaseItemObsList = FXCollections.observableArrayList();

    public RotateItem(int rowNum, String kitName, String partNum, int pmQty) {
        if (kitName == null)
            kitName = "";

        this.kitName = new SimpleStringProperty(kitName);
        this.partNumber = new SimpleStringProperty(partNum);
        this.pmQty = new SimpleIntegerProperty(pmQty);
        this.myQty = new SimpleIntegerProperty(0);


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

    public void setDuplicate(List<RotateItem> list) {
        this.isDuplicate = true;
        this.duplicateRotateItems = list;
        duplicateRotateItems.add(this);
    }

    public void addDuplicate(RotateItem item)
    {
        isDuplicate = true;
        duplicateRotateItems = new ArrayList<>();
        duplicateRotateItems.add(this);

        item.setDuplicate(duplicateRotateItems);
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

    public void addStockItem(StockItem item) {

        log.debug(String.format("Add stock: %s, %s, %s, %d, %d, %d",
                item.getKitName(), item.getPartNumber(), item.getPo(), item.getStockQty(), item.getDc(),
                item.getMyQty()));

        addMyQty(item.getMyQty());

        item.setRotateItem(this);

        stockItems.put(item.getPo(), item);
        stockItemObsList.add(item);
    }

    public ObservableList<StockItem> getStockItemObsList() {
        return stockItemObsList;
    }

    public int getMyQty() {
        return myQty.get();
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

    public void addNoneStockPurchase(PurchaseItem purchaseItem) {
        noneStockPurchaseItemObsList.add(purchaseItem);
    }
}
