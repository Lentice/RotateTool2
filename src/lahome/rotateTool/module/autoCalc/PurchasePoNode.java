package lahome.rotateTool.module.autoCalc;

import gnu.trove.map.hash.THashMap;
import lahome.rotateTool.Util.UndoManager;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateItem;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by LenticeTsai on 2017/5/25.
 */
public class PurchasePoNode {
    private static final String REMARK_HEAD = "實退 ";
    private static UndoManager undoManager;

    private String po = null;
    private int maxSet = 0;
    private int applySet = 0;
    private Object[] kitPurchaseItemList;

    private PurchasePoNode(int kitCount) {
        kitPurchaseItemList = new Object[kitCount];
        undoManager = UndoManager.getInstance();
    }

    public void addPurchaseItem(int partIndex, PurchaseItem purchaseItem) {
        if (po == null)
            po = purchaseItem.getPo();

        //noinspection unchecked
        List<PurchaseItem> purchaseItemList = (List<PurchaseItem>) kitPurchaseItemList[partIndex];
        if (purchaseItemList == null) {
            purchaseItemList = new ArrayList<>();
            kitPurchaseItemList[partIndex] = purchaseItemList;
        }

        purchaseItemList.add(purchaseItem);
    }

    public void addApplySet(int set, String po, int[] ratio) {
        applySet += set;
        for (int i = 0; i < ratio.length; i++) {
            int expectQty = set * ratio[i];
            //noinspection unchecked
            List<PurchaseItem> purchaseItemList = (List<PurchaseItem>) kitPurchaseItemList[i];
            for (PurchaseItem purchaseItem: purchaseItemList) {
                if (expectQty <= 0)
                    break;

                int grQty = purchaseItem.getGrQty();
                int oldApplyQty = purchaseItem.getApplyQty();
                int applyQty = Math.min(expectQty, grQty - oldApplyQty);

                undoManager.appendInput(purchaseItem.applyQtyProperty(), applyQty + oldApplyQty);

                String remark = purchaseItem.getRemark();
                if (remark.isEmpty())
                    remark += REMARK_HEAD + po;
                else
                    remark += ", " + po;

                undoManager.appendInput(purchaseItem.remarkProperty(), remark);

                expectQty -= applyQty;
            }
        }
    }

    public void clearRedundantRemark() {
        for (int i = 0; i < kitPurchaseItemList.length; i++) {
            List<PurchaseItem> purchaseItemList = getKitPurchaseItemList(i);
            if (purchaseItemList == null)
                continue;

            for (PurchaseItem purchaseItem: purchaseItemList) {
                if (purchaseItem.getRemark().equals(REMARK_HEAD + purchaseItem.getPo())) {
                    undoManager.appendInput(purchaseItem.remarkProperty(), "");
                }
            }
        }
    }

    private void postProcess(int[] ratio) {
        calcMaxSet(ratio);
    }

    private void calcMaxSet(int[] ratio) {
        maxSet = Integer.MAX_VALUE;
        for (int i = 0; i < ratio.length; i++) {
            //noinspection unchecked
            List<PurchaseItem> purchaseItemList = (List<PurchaseItem>) kitPurchaseItemList[i];
            int totalGrQty = 0;
            for (PurchaseItem purchaseItem : purchaseItemList) {
                totalGrQty += purchaseItem.getGrQty();
            }
            maxSet = Math.min(maxSet, totalGrQty / ratio[i]);
        }
    }

    public String getPo() {
        return po;
    }

    public int getMaxSet() {
        return maxSet;
    }

    public int getApplySet() {
        return applySet;
    }

    public void setApplySet(int applySet) {
        this.applySet = applySet;
    }

    public int getRemainderSet() {
        return maxSet - applySet;
    }

    public List<PurchaseItem> getKitPurchaseItemList(int partIndex) {
        //noinspection unchecked
        return (List<PurchaseItem>) kitPurchaseItemList[partIndex];
    }

    public static List<PurchasePoNode> getPoNodeList(RotateItem[] kit, int[] ratio) {
        assert kit.length == ratio.length;

        THashMap<String, PurchasePoNode> poNodeMap = new THashMap<>();

        for (int i = 0; i < kit.length; i++) {
            for (PurchaseItem purchaseItem : kit[i].getStAndNoneStPurchaseItemList()) {
                String key = purchaseItem.getPo();
                PurchasePoNode poNode = poNodeMap.get(key);
                if (poNode == null) {
                    poNode = new PurchasePoNode(kit.length);
                    poNodeMap.put(key, poNode);
                }

                poNode.addPurchaseItem(i, purchaseItem);
            }
        }

        List<PurchasePoNode> poNodeList = new ArrayList<>(poNodeMap.values());
        for (PurchasePoNode poNode : poNodeList) {
            poNode.postProcess(ratio);
        }

        return poNodeList;
    }

}
