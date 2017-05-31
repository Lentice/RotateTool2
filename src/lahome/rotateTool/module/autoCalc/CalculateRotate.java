package lahome.rotateTool.module.autoCalc;

import lahome.rotateTool.Util.UndoManager;
import lahome.rotateTool.module.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class CalculateRotate {

    private RotateCollection collection;
    private int dcMinLimit;
    private UndoManager undoManager;

    private int[] ratio;
    private RotateItem[] kit;

    private List<StockPoNode> stockPoNodeList;
    private List<PurchasePoNode> purchasePoNodeList;

    public CalculateRotate(RotateCollection collection, int dcMinLimit) {
        this.collection = collection;
        this.dcMinLimit = dcMinLimit;
        undoManager = UndoManager.getInstance();
    }

    public void calculateAll() {

        undoManager.newInput();

        clearAllForAutoCalc();
        for (RotateItem rotateItem : collection.getRotateObsList()) {
            processOneRotate(rotateItem);
        }

        clean();
    }

    public void calculateOneRotate(RotateItem rotateItem) {

        undoManager.newInput();

        clearOneRotateForAutoCalc(rotateItem);
        processOneRotate(rotateItem);

        clean();
    }

    private void clean() {
        ratio = null;
        kit = null;

        stockPoNodeList = null;
        purchasePoNodeList = null;

        undoManager = null;
        collection = null;
    }

    private boolean isValidKit() {
        for (RotateItem rotateItem : kit) {
            if (!rotateItem.isRotateValid())
                return false;
        }

        return true;
    }

    private void processOneRotate(RotateItem rotateItem) {
        if (rotateItem.getBacklog() != 0) {
            undoManager.appendInput(rotateItem.remarkProperty(), "Backlog 不為 0");
            return;
        }

        if (rotateItem.isDuplicate()) {
            undoManager.appendInput(rotateItem.remarkProperty(), "重複項目，請記得處理");
            return;
        }

        if (!rotateItem.isRotateValid())
            return;

        kit = getKitItems(rotateItem);
        if (kit == null)
            return;

        if (!isValidKit())
            return;

        processStockItems();
    }

    private void processStockItems() {
        ratio = getRatio();
        assert ratio != null;

        int remainderSetTotal = getTargetSet();

        stockPoNodeList = StockPoNode.getPoNodeList(kit, ratio, dcMinLimit);
        purchasePoNodeList = PurchasePoNode.getPoNodeList(kit, ratio);

        // Sort by DC. If equal then sort by MaxSet.
        stockPoNodeList.sort((o1, o2) -> {
            int val = o1.getDc() - o2.getDc();
            if (val != 0)
                return val;

            return -(o1.getMaxSet() - o2.getMaxSet());
        });

        for (StockPoNode stockPoNode : stockPoNodeList) {
            if (remainderSetTotal <= 0)
                break;

            int set = Math.min(remainderSetTotal, stockPoNode.getMaxSet());
            stockPoNode.setTargetSet(set);
            remainderSetTotal -= set;
        }

        for (StockPoNode stockPoNode : stockPoNodeList) {
            if (stockPoNode.getMatchedPo() == kit.length) {
                fillApplyQty(stockPoNode, true);
            }
        }

        stockPoNodeList.sort(Comparator.comparingInt(StockPoNode::getRemainderSet).reversed());
        for (StockPoNode stockPoNode : stockPoNodeList) {
            if (stockPoNode.getRemainderSet() <= 0)
                continue;

            purchasePoNodeList.sort(Comparator.comparingInt(PurchasePoNode::getRemainderSet).reversed());
            fillApplyQty(stockPoNode, false);
        }

        clearRedundantRemark();
    }

    private void clearRedundantRemark() {
        for (StockPoNode stockPoNode: stockPoNodeList) {
            stockPoNode.clearRedundantRemark();
        }

        for (PurchasePoNode purchasePoNode: purchasePoNodeList) {
            purchasePoNode.clearRedundantRemark();
        }
    }

    private void fillApplyQty(StockPoNode stockPoNode, boolean matchedPoOnly) {
        for (PurchasePoNode purchasePoNode : purchasePoNodeList) {
            if (stockPoNode.getRemainderSet() <= 0)
                return;

            if (matchedPoOnly && !purchasePoNode.getPo().equals(stockPoNode.getPo()))
                continue;

            int set = Math.min(stockPoNode.getRemainderSet(), purchasePoNode.getRemainderSet());
            stockPoNode.addApplySet(set, purchasePoNode.getPo(), ratio);
            purchasePoNode.addApplySet(set, stockPoNode.getPo(), ratio);
        }
    }


    private int getTargetSet() {
        int targetSet = Integer.MAX_VALUE;

        for (int i = 0; i < kit.length; i++) {
            targetSet = Math.min(targetSet, kit[i].getPmQty() / ratio[i]);
        }

        for (int i = 0; i < kit.length; i++) {
            int grQty = 0;
            for (PurchaseItem purchaseItem : kit[i].getStAndNoneStPurchaseItemList()) {
                grQty += purchaseItem.getGrQty();
            }

            targetSet = Math.min(targetSet, grQty / ratio[i]);
        }

        for (int i = 0; i < kit.length; i++) {
            int stockQty = 0;
            for (StockItem stockItem : kit[i].getStockItemObsList()) {
                stockQty += stockItem.getStockQty();
            }

            targetSet = Math.min(targetSet, stockQty / ratio[i]);
        }

        return targetSet;
    }

    private void clearRotateItem(RotateItem rotateItem) {
        undoManager.appendInput(rotateItem.remarkProperty(), "");
    }

    private void clearStockItem(StockItem stockItem) {
        stockItem.setAutoCalculated(false);

        undoManager.appendInput(stockItem.applyQtyProperty(), 0);
        undoManager.appendInput(stockItem.remarkProperty(), "");
    }

    private void clearPurchaseItem(PurchaseItem purchaseItem) {
        undoManager.appendInput(purchaseItem.applyQtyProperty(), 0);
        undoManager.appendInput(purchaseItem.remarkProperty(), "");
    }

    private void clearAllForAutoCalc() {
        for (RotateItem rotateItem : collection.getRotateObsList()) {
            clearRotateItem(rotateItem);
        }

        for (StockItem stockItem : collection.getStockItemList()) {
            clearStockItem(stockItem);
        }

        for (PurchaseItem purchaseItem : collection.getPurchaseItemList()) {
            clearPurchaseItem(purchaseItem);
        }
    }

    private void clearOneRotateForAutoCalc(RotateItem rotateItem) {

        List<RotateItem> rotateItemList = null;
        List<StockItem> stockItemList = null;
        List<PurchaseItem> purchaseItemList = null;

        for (RotateItem item : collection.getRotateObsList()) {
            if (rotateItem.isKit() && item.isKit() && item.getKitName().equals(rotateItem.getKitName())) {
                KitNode kitNode = item.getKitNode();

                rotateItemList = kitNode.getPartsList();
                stockItemList = kitNode.getStockItemList();
                purchaseItemList = kitNode.getStAndNoneStPurchaseItemList();

                break;
            } else if (!rotateItem.isKit() && !item.isKit()) {
                if (item.getPartNumber().equals(rotateItem.getPartNumber())) {
                    rotateItemList = new ArrayList<>(2);
                    rotateItemList.add(rotateItem);

                    stockItemList = rotateItem.getStockItemObsList();
                    purchaseItemList = rotateItem.getStAndNoneStPurchaseItemList();
                    break;
                }
            }
        }

        if (rotateItemList != null) {
            for (RotateItem kitItem : rotateItemList) {
                clearRotateItem(kitItem);
            }
        }

        if (stockItemList != null) {
            for (StockItem stockItem : stockItemList) {
                clearStockItem(stockItem);
            }
        }

        if (purchaseItemList != null) {
            for (PurchaseItem purchaseItem : purchaseItemList) {
                clearPurchaseItem(purchaseItem);
            }
        }
    }

    private int[] getRatio() {
        int[] ratio = new int[kit.length];
        for (int i = 0; i < kit.length; i++) {
            if (kit[i].isKit())
                ratio[i] = kit[i].getRatio();
            else
                ratio[i] = 1;

            if (ratio[i] == 0)
                return null;
        }
        return ratio;
    }

    private RotateItem[] getKitItems(RotateItem rotateItem) {
        if (rotateItem.isKit()) {
            if (rotateItem.isFirstPartOfKit()) {
                List<RotateItem> items = rotateItem.getKitNode().getPartsList();
                RotateItem[] kit = new RotateItem[items.size()];
                int index = 0;
                for (RotateItem item : items) {
                    kit[index++] = item;
                }
                return kit;
            } else {
                return null;
            }
        } else {
            RotateItem[] kit = new RotateItem[1];
            kit[0] = rotateItem;
            return kit;
        }
    }

}
