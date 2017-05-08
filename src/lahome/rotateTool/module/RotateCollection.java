package lahome.rotateTool.module;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class RotateCollection {
    private static final Logger log = LogManager.getLogger(RotateCollection.class.getName());

    private HashMap<String, KitNode> kitNodeMap = new HashMap<>();
    private HashMap<String, RotateItem> rotateItemMap = new HashMap<>();

    private ObservableList<RotateItem> rotateObsList = FXCollections.observableArrayList();
    private List<StockItem> stockItemList = new ArrayList<>();
    private List<PurchaseItem> purchaseItemList = new ArrayList<>();

    private String getRotateKey(String kitName, String partNum) {
        return kitName.replaceAll("-", "") + partNum.replaceAll("-", "");
    }

    public void addRotate(RotateItem item) {

        log.debug(String.format("Add rotate: %s, %s, %d, %s",
                item.getKitName(), item.getPartNumber(), item.getPmQty(), item.isKit() ? "K" : "S"));

        String key = getRotateKey(item.getKitName(), item.getPartNumber());
        RotateItem rotateItem = rotateItemMap.get(key);
        if (rotateItem != null) {
            rotateItem.addDuplicate(item);
        } else {
            rotateItemMap.put(key, item);
        }

        rotateObsList.add(item);

        if (item.isKit() && !item.isDuplicate()) {
            KitNode kitNode = kitNodeMap.computeIfAbsent(item.getKitName(), k -> new KitNode());
            kitNode.addPart(item);
        }
    }

    public void addStock(String kitName, String partNum, StockItem item) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItemMap.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed stock item %s, %s", kitName, partNum));
            item.clear();
            return;
        }

        stockItemList.add(item);
        rotateItem.addStockItem(item);
    }

    public void addPurchase(String kitName, String partNum, PurchaseItem item) {

        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItemMap.get(key);
        if (rotateItem == null) {
            log.trace(String.format("Ignore no needed purchase item %s, %s", kitName, partNum));
            item.clear();
            return;
        }

        StockItem stockitem = rotateItem.getStockByPo(item.getPo());
        if (stockitem == null) {
            rotateItem.addNoneStockPurchase(item);
        } else {
            stockitem.addPurchaseItem(item);
        }
        purchaseItemList.add(item);
    }

    public boolean belongToRotateItem(String kitName, String partNum) {
        String key = getRotateKey(kitName, partNum);
        RotateItem rotateItem = rotateItemMap.get(key);
        return (rotateItem != null);
    }

    public ObservableList<RotateItem> getRotateObsList() {
        return rotateObsList;
    }

    public List<StockItem> getStockItemList() {
        return stockItemList;
    }

    public List<PurchaseItem> getPurchaseItemList() {
        return purchaseItemList;
    }

    public HashMap<String, KitNode> getKitNodeMap() {
        return kitNodeMap;
    }

}
