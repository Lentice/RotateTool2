package lahome.rotateTool.module;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;

import java.util.HashMap;

public class RotateCollection {
    private static final Logger log = LogManager.getLogger(RotateCollection.class.getName());

    private HashMap<String, RotateItem> rotateItems = new HashMap<>();

    private ObservableList<RotateItem> rotateObsList = FXCollections.observableArrayList();

    private int partCount = 0;

    public RotateCollection() {
    }

    private String getRotateKey(String kitName, String partNum) {
        return kitName.replaceAll("-", "") + partNum.replaceAll("-", "");
    }

    public void addRotate(RotateItem item) {

        log.debug(String.format("Add rotate: %s, %s, %d, %s",
                item.getKitName(), item.getPartNumber(), item.getPmQty(), item.isKit()? "K" : "S"));

        String key = getRotateKey(item.getKitName(), item.getPartNumber());
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem != null) {
            rotateItem.addDuplicate(item);
        }

        rotateItems.put(key, item);
        rotateObsList.add(item);

        partCount++;
    }

    public void addStock(Row row, String kitName, String partNum, String po,
                         int stockQty, int dc, int myQty) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed stock item %s, %s", kitName, partNum));
            return;
        }

        StockItem item = new StockItem(row, kitName, partNum, po, stockQty, dc, myQty);
        rotateItem.addStockItem(item);
    }

    public void addPurchase(Row row, String kitName, String partNum, String po, String date, int grQty) {

        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed purchase item %s, %s", kitName, partNum));
            return;
        }

        PurchaseItem purchaseItem = new PurchaseItem(row, kitName, partNum, po, date, grQty);
        StockItem stockitem = rotateItem.getStock(po);
        if (stockitem == null) {
            rotateItem.addNoneStockPurchase(purchaseItem);
        } else {
            stockitem.addPurchaseItem(purchaseItem);
        }
    }

    public boolean isRotateItem(String kitName, String partNum) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        return (rotateItem != null);
    }

    public int getPartCount() {
        return partCount;
    }

    public ObservableList<RotateItem> getRotateObsList() {
        return rotateObsList;
    }
}
