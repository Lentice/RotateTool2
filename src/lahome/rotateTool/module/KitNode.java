package lahome.rotateTool.module;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

public class KitNode {

    private ObservableList<RotateItem> PartsList = FXCollections.observableArrayList();

    public void addPart(RotateItem item) {
        if (item.isDuplicate())
            return;

        PartsList.add(item);
        int partCount = PartsList.size();
        if (partCount == 1) {
            item.setFirstPartOfKit(true);
            item.updateApplySet();
            item.applySetListenerStart();
        }

        item.setSerialNo(String.valueOf((char) (partCount - 1 + 'A')));
        item.setKitNode(this);
    }

    public RotateItem getPart(String partNum) {
        for (RotateItem rotateItem : PartsList) {
            if (partNum.equals(rotateItem.getPartNumber())) {
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
                if (stockItem.isMainStockItem() && po.equals(stockItem.getPo())) {
                    list.addAll(stockItem.getPurchaseItemList());
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
                    list.addAll(stockItem.getPurchaseItemList());
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
