package lahome.rotateTool.module;

import javafx.beans.property.IntegerProperty;
import javafx.beans.property.SimpleIntegerProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

/**
 * Created by LenticeTsai on 2017/5/5.
 */
public class KitNode {

    private ObservableList<RotateItem> PartsList = FXCollections.observableArrayList();

    public void addPart(RotateItem item) {
        if (item.isDuplicate())
            return;

        item.setSerialNo(String.valueOf((char) (PartsList.size() + 'A')));

        PartsList.add(item);
        item.setKitNode(this);
    }

    public RotateItem getPart(String partNum) {
        for (RotateItem rotateItem : PartsList) {
            if (partNum.compareTo(rotateItem.getPartNumber()) == 0) {
                return rotateItem;
            }
        }
        return null;
    }

    public ObservableList<RotateItem> getPartsList() {
        return PartsList;
    }

    public ObservableList<StockItem> getStockItemList() {
        ObservableList<StockItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            list.addAll(rotateItem.getStockItemObsList());
        }
        list.sort(StockItem::compareTo);

        return list;
    }

    public ObservableList<PurchaseItem> getPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            list.addAll(rotateItem.getPurchaseItemList());
        }

        list.sort(PurchaseItem::compareTo);


        return list;
    }

    public ObservableList<PurchaseItem> getPurchaseItemListWithPo(String po) {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            for (StockItem stockItem : rotateItem.getStockItemObsList()) {
                if (stockItem.isMainStockItem() && po.compareTo(stockItem.getPo()) == 0) {
                    list.addAll(stockItem.getPurchaseItems());
                }
            }
        }

        list.sort(PurchaseItem::compareTo);


        return list;
    }

    public ObservableList<PurchaseItem> getNoneStPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            list.addAll(rotateItem.getNoneStockPurchaseItemObsList());
        }

        list.sort(PurchaseItem::compareTo);

        return list;
    }

    public ObservableList<PurchaseItem> getStAndNoneStPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            for (StockItem stockItem : rotateItem.getStockItemObsList()) {
                if (stockItem.isMainStockItem()) {
                    list.addAll(stockItem.getPurchaseItems());
                }
            }
        }
        for (RotateItem rotateItem : PartsList) {
            list.addAll(rotateItem.getNoneStockPurchaseItemObsList());
        }

        list.sort(PurchaseItem::compareTo);

        return list;
    }

    public int pmQtyProperty() {
        int pmQtyTotal = 0;
        for (RotateItem rotateItem : PartsList) {
            pmQtyTotal += rotateItem.getPmQty();
        }
        return pmQtyTotal;
    }
}
