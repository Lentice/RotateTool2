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

    private IntegerProperty purchaseAllApSetTotal = new SimpleIntegerProperty(0);

    private IntegerProperty stockQtyTotal = new SimpleIntegerProperty(0);

    public KitNode() {

    }

    public void addPart(RotateItem item) {
        if (item.isDuplicate())
            return;

        item.setSerialNo(String.valueOf((char) (PartsList.size() + 'A')));

        PartsList.add(item);
        item.setKitNode(this);
    }

    public IntegerProperty purchaseAllApSetTotalProperty() {
        return purchaseAllApSetTotal;
    }

    public void addPurchaseAllApSetTotal(int set) {
        purchaseAllApSetTotal.set(purchaseAllApSetTotal.intValue() + set);
    }

    public IntegerProperty stockQtyTotalProperty() {
        return stockQtyTotal;
    }

    public void addStockQtyTotal(int qty) {
        stockQtyTotal.set(stockQtyTotal.intValue() + qty);
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

        // sort by Gr Date by default
        list.sort((o1, o2) -> {
            if (o2.getEarliestGrDate().isEmpty()) {
                return -1;
            } else if (o1.getEarliestGrDate().isEmpty()) {
                return 1;
            } else {
                int value = o1.getEarliestGrDate().compareTo(o2.getEarliestGrDate());
                if (value != 0)
                    return value;

                value = o1.getPo().compareTo(o2.getPo());
                if (value != 0)
                    return value;

                return o1.getRotateItem().getPartNumber().compareTo(o2.getRotateItem().getPartNumber());
            }
        });

        return list;
    }

    public ObservableList<PurchaseItem> getPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            for (StockItem stockItem : rotateItem.getStockItemObsList()) {
                if (stockItem.isMainStockItem()) {
                    list.addAll(stockItem.getPurchaseItems());
                }
            }
        }

        // sort by Gr Date by default
        list.sort((o1, o2) -> {
            int value = o1.getGrDate().compareTo(o2.getGrDate());
            if (value != 0)
                return value;

            value = o1.getPo().compareTo(o2.getPo());
            if (value != 0)
                return value;

            return o1.getRotateItem().getPartNumber().compareTo(o2.getRotateItem().getPartNumber());
        });


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

        // sort by Gr Date by default
        list.sort((o1, o2) -> {
            int value = o1.getGrDate().compareTo(o2.getGrDate());
            if (value != 0)
                return value;

            value = o1.getPo().compareTo(o2.getPo());
            if (value != 0)
                return value;

            return o1.getRotateItem().getPartNumber().compareTo(o2.getRotateItem().getPartNumber());
        });


        return list;
    }

    public ObservableList<PurchaseItem> getNoneStPurchaseItemList() {
        ObservableList<PurchaseItem> list = FXCollections.observableArrayList();
        for (RotateItem rotateItem : PartsList) {
            list.addAll(rotateItem.getNoneStockPurchaseItemObsList());
        }

        // sort by Gr Date by default
        list.sort((o1, o2) -> {
            int value = o1.getGrDate().compareTo(o2.getGrDate());
            if (value != 0)
                return value;

            value = o1.getPo().compareTo(o2.getPo());
            if (value != 0)
                return value;

            return o1.getRotateItem().getPartNumber().compareTo(o2.getRotateItem().getPartNumber());
        });

        return list;
    }

    public IntegerProperty getPmQtyProperty() {
        int pmQtyTotal = 0;
        for (RotateItem rotateItem : PartsList) {
            pmQtyTotal += rotateItem.getPmQty();
        }
        return new SimpleIntegerProperty(pmQtyTotal);
    }
}
