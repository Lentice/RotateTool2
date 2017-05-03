package lahome.rotateTool.module;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;

import java.util.Comparator;
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
                item.getKitName(), item.getPartNumber(), item.getPmQty(), item.isKit() ? "K" : "S"));

        String key = getRotateKey(item.getKitName(), item.getPartNumber());
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem != null) {
            rotateItem.addDuplicate(item);
        } else {
            rotateItems.put(key, item);
        }

        rotateObsList.add(item);

        partCount++;
    }

    public void addStock(String kitName, String partNum, StockItem item) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed stock item %s, %s", kitName, partNum));
            item.close();
            return;
        }

        rotateItem.addStockItem(item);
    }

    public void addPurchase(String kitName, String partNum, PurchaseItem item) {

        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed purchase item %s, %s", kitName, partNum));
            item.close();
            return;
        }

        StockItem stockitem = rotateItem.getStock(item.getPo());
        if (stockitem == null) {
            rotateItem.addNoneStockPurchase(item);
        } else {
            stockitem.addPurchaseItem(item);
        }
    }

    public boolean isRotateItem(String kitName, String partNum) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItems.get(key);
        return (rotateItem != null);
    }

    public ObservableList<RotateItem> getRotateObsList() {
        return rotateObsList;
    }
}
