package lahome.rotateTool.Util;

import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Administrator on 2017/5/14.
 */
public class CalculateRotate {

    private static UndoManager undoManager;
    private static int targetSet;
    private static int[] ratio;
    private static int[] applyQty;
    private static RotateItem[] kit;
    private static List<PurchaseItem> applyNoneStPurchase;

    public static void doCalculate(RotateCollection collection, int dcMinLimit) {

        undoManager = UndoManager.getInstance();
        undoManager.newInput();

        clearForAutoCalc(collection);
        for (RotateItem rotateItem : collection.getRotateObsList()) {
            if (rotateItem.isDuplicate())
                continue;

            if (!rotateItem.isRotateValid())
                continue;

            kit = getKitItems(rotateItem);
            if (kit == null)
                continue;

            if (!isValidKit())
                continue;

            processStockItems(dcMinLimit);
        }
    }

    private static boolean isValidKit() {
        for (RotateItem rotateItem : kit) {
            if (!rotateItem.isRotateValid())
                return false;

            if (rotateItem.getBacklog() != 0)
                return false;
        }

        return true;
    }

    private static void processStockItems(int dcMinLimit) {

        ratio = getRatio();
        applyQty = new int[kit.length];

        assert ratio != null;
        targetSet = kit[0].getPmQty() / ratio[0];

        List<StockItem> stockList = getAllStockList(dcMinLimit);
        for (StockItem stockItem : stockList) {
            if (stockItem.isAutoCalculated())
                continue;

            processApplyQty(stockItem);

            if ((applyQty[0] / ratio[0]) >= targetSet) {
                break;
            }
        }
    }

    private static void processApplyQty(StockItem stockItem) {
        applyNoneStPurchase = new ArrayList<>();
        boolean useNoneStPurchase = (stockItem.getPurchaseItemList().size() == 0);

//        if ("4501257283".equals(stockItem.getPo())) {
//            int i = 1;
//        }

        int stockApplySet = getMinApplySetFromStock(stockItem.getPo());
        int applySet = Math.min(stockApplySet, targetSet - ((applyQty[0]) / ratio[0]));
        if (applySet == 0)
            return;

        int purchaseApplySet = (useNoneStPurchase) ?
                getMinApplySetFromNoneStPurchase(applySet) :
                getMinApplySetFromPurchase(stockItem.getPo());

        applySet = Math.min(applySet, purchaseApplySet);
        if (applySet == 0)
            return;


        applyNoneStPurchase.sort((o1, o2) -> o2.getApplyQty() - o1.getApplyQty());

        fillApplyQtyToStock(stockItem.getPo(), applySet, useNoneStPurchase);
        if (useNoneStPurchase) {
            fillApplyQtyToNoneStPurchase(stockItem.getPo(), applySet);
        } else {
            fillApplyQtyToPurchase(stockItem.getPo(), applySet);
        }

        for (int i = 0; i < kit.length; i++) {
            applyQty[i] += applySet * ratio[i];
        }
    }

    private static void fillApplyQtyToPurchase(String po, int applySet) {
        for (int i = 0; i < kit.length; i++) {
            int expectQty = applySet * ratio[i];
            for (PurchaseItem purchaseItem : kit[i].getPurchaseItemList()) {
                if (purchaseItem.getPo().equals(po)) {
                    int grQty = purchaseItem.getGrQty();
                    int oldApplyQty = purchaseItem.getApplyQty();
                    int applyQty = Math.min(expectQty, grQty - oldApplyQty);

                    undoManager.appendInput(purchaseItem.applyQtyProperty(),
                            applyQty + oldApplyQty);

                    expectQty -= applyQty;
                    if (expectQty <= 0)
                        break;
                }
            }
        }
    }

    private static void fillApplyQtyToNoneStPurchase(String po, int applySet) {
        for (int i = 0; i < kit.length; i++) {
            int expectQty = applySet * ratio[i];
            for (PurchaseItem applyPurchase : applyNoneStPurchase) {
                String applyPo = applyPurchase.getPo();
                for (PurchaseItem purchaseItem : kit[i].getNoneStockPurchaseItemObsList()) {
                    if (purchaseItem.getPo().equals(applyPo)) {
                        int grQty = purchaseItem.getGrQty();
                        int oldApplyQty = purchaseItem.getApplyQty();
                        int applyQty = Math.min(expectQty, grQty - oldApplyQty);

                        undoManager.appendInput(purchaseItem.applyQtyProperty(),
                                applyQty + oldApplyQty);

                        String remark = purchaseItem.getRemark();
                        if (remark.isEmpty())
                            remark += "實退 " + po;
                        else
                            remark += ", " + po;

                        undoManager.appendInput(purchaseItem.remarkProperty(),
                                remark);

                        expectQty -= applyQty;
                        if (expectQty <= 0)
                            break;
                    }
                }

                if (expectQty <= 0)
                    break;
            }
        }
    }

    private static void fillApplyQtyToStock(String po, int applySet, boolean useNoneStPurchase) {
        StringBuilder remark = new StringBuilder("apply by PO#");
        if (useNoneStPurchase) {
            for (PurchaseItem applyPurchase : applyNoneStPurchase) {
                String applyPo = applyPurchase.getPo();
                remark.append(applyPo).append(", ");
            }
            remark.deleteCharAt(remark.length() - 2);
        }

        for (int i = 0; i < kit.length; i++) {
            int expectQty = applySet * ratio[i];
            for (StockItem stockItem : kit[i].getStockItemObsList()) {
                if (stockItem.getPo().equals(po)) {
                    int stockQty = stockItem.getStockQty();
                    int oldApplyQty = stockItem.getApplyQty();
                    int applyQty = Math.min(expectQty, stockQty - oldApplyQty);

                    undoManager.appendInput(stockItem.applyQtyProperty(),
                            applyQty + oldApplyQty);

                    if (useNoneStPurchase) {
                        undoManager.appendInput(stockItem.remarkProperty(),
                                remark.toString());
                    }

                    stockItem.setAutoCalculated(true);

                    expectQty -= applyQty;
                    if (expectQty <= 0) {
                        break;
                    }
                }
            }
        }
    }

    private static int getMinApplySetFromPurchase(String po) {
        int[] grQty = new int[kit.length];

        for (int i = 0; i < kit.length; i++) {
            for (PurchaseItem purchaseItem : kit[i].getPurchaseItemList()) {
                if (purchaseItem.getPo().equals(po)) {
                    grQty[i] += (purchaseItem.getGrQty() - purchaseItem.getApplyQty());
                }
            }
        }

        int minApplySet = grQty[0] / ratio[0];
        for (int i = 1; i < kit.length; i++) {
            int set = grQty[i] / ratio[i];
            if (set < minApplySet)
                minApplySet = set;
        }

        return minApplySet;
    }


    private static int getMinApplySetFromNoneStPurchase(int expectSet) {

        int applySetTotal = 0;
        String selectedPo = "";

        List<PurchaseItem> purchaseList = new ArrayList<>(kit[0].getNoneStockPurchaseItemObsList());
        purchaseList.sort((o1, o2) -> (o2.getGrQty() - o2.getApplyQty()) - (o1.getGrQty() - o1.getApplyQty()));

        for (PurchaseItem poItem : purchaseList) {
            if (poItem.getPo().equals(selectedPo))
                continue;

            selectedPo = poItem.getPo();
            int[] grQty = new int[kit.length];
            for (int i = 0; i < kit.length; i++) {
                for (PurchaseItem purchaseItem : kit[i].getNoneStockPurchaseItemObsList()) {
                    if (purchaseItem.getPo().equals(selectedPo))
                        grQty[i] += (purchaseItem.getGrQty() - purchaseItem.getApplyQty());
                }
            }

            int minApplySet = grQty[0] / ratio[0];
            for (int i = 1; i < kit.length; i++) {
                int set = grQty[i] / ratio[i];
                if (set < minApplySet)
                    minApplySet = set;
            }

            if (minApplySet > 0) {
                applySetTotal += minApplySet;
                applyNoneStPurchase.add(poItem);
            }

            if (applySetTotal >= expectSet)
                return expectSet;
        }

        return applySetTotal;
    }

    private static int getMinApplySetFromStock(String po) {
        int[] stockQty = new int[kit.length];
        for (int i = 0; i < kit.length; i++) {
            for (StockItem stockItem : kit[i].getStockItemObsList()) {
                if (stockItem.getPo().equals(po)) {
                    stockQty[i] += stockItem.getStockQty() - stockItem.getApplyQty();
                }
            }
        }

        int minApplySet = stockQty[0] / ratio[0];
        for (int i = 1; i < kit.length; i++) {
            int set = stockQty[i] / ratio[i];
            if (set < minApplySet)
                minApplySet = set;
        }
        return minApplySet;
    }

    private static void clearForAutoCalc(RotateCollection collection) {
        for (StockItem stockItem : collection.getStockItemList()) {
            undoManager.appendInput(stockItem.applyQtyProperty(), 0);
            undoManager.appendInput(stockItem.remarkProperty(), "");


            stockItem.setAutoCalculated(false);
        }

        for (PurchaseItem purchaseItem : collection.getPurchaseItemList()) {
            undoManager.appendInput(purchaseItem.applyQtyProperty(), 0);
            undoManager.appendInput(purchaseItem.remarkProperty(), "");
        }
    }

    private static List<StockItem> getAllStockList(int dcMin) {
        List<StockItem> list = new ArrayList<>();

        for (RotateItem rotateItem : kit) {
            for (StockItem stockItem : rotateItem.getStockItemObsList()) {
                if (stockItem.getDc() >= dcMin) {
                    list.add(stockItem);
                }
            }
        }

        list.sort((o1, o2) -> {
            int dc = o2.getDc() - o1.getDc();
            if (dc != 0)
                return -dc;

            return (o2.getStockQty() - o1.getStockQty());
        });

        return list;
    }

    private static int[] getRatio() {
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

    private static RotateItem[] getKitItems(RotateItem rotateItem) {
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
