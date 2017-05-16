package lahome.rotateTool.module;

import gnu.trove.map.hash.THashMap;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.commons.lang3.StringUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.ArrayList;
import java.util.List;

public class RotateCollection {
    private static final Logger log = LogManager.getLogger(RotateCollection.class.getName());

    private THashMap<String, KitNode> kitNodeMap = new THashMap<>();
    private THashMap<String, RotateItem> rotateItemMap = new THashMap<>();

    private ObservableList<RotateItem> rotateObsList = FXCollections.observableArrayList();
    private List<StockItem> stockItemList = new ArrayList<>();
    private List<PurchaseItem> purchaseItemList = new ArrayList<>();

    int maxDc = 0;

    private String getRotateKey(String kitName, String partNum) {
        return StringUtils.remove(kitName + partNum, '-');
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

    public void addStock(RotateItem rotateItem, StockItem item) {
        assert rotateItem != null;

        stockItemList.add(item);
        rotateItem.addStockItem(item);

        maxDc = Math.max(maxDc, item.getDc());
    }

    public void addPurchase(RotateItem rotateItem, PurchaseItem item) {
        assert  rotateItem != null;

        StockItem stockitem = rotateItem.getStockByPo(item.getPo());
        if (stockitem == null) {
            rotateItem.addNoneStockPurchase(item);
        } else {
            stockitem.addPurchaseItem(item);
        }
        purchaseItemList.add(item);
    }

    public RotateItem getRotateItem(String kitName, String partNum) {
        String key = getRotateKey(kitName, partNum);
        return rotateItemMap.get(key);
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

    public THashMap<String, KitNode> getKitNodeMap() {
        return kitNodeMap;
    }

    public int getMaxDc() {
        return maxDc;
    }
}
