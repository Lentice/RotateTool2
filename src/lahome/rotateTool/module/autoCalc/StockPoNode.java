package lahome.rotateTool.module.autoCalc;

import gnu.trove.map.hash.THashMap;
import lahome.rotateTool.Util.UndoManager;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by LenticeTsai on 2017/5/25.
 */
public class StockPoNode {
    private static final String REMARK_HEAD = "apply by PO#";

    private static UndoManager undoManager;

    private String po = null;

    private int maxSet = 0;
    private int targetSet = 0;
    private int applySet = 0;
    private int dc = Integer.MAX_VALUE;

    private int matchedPo = 0; // how many part have matched PO in purchase

    private Object[] kitStockItemList;

    public StockPoNode(int kitCount) {
        kitStockItemList = new Object[kitCount];
        undoManager = UndoManager.getInstance();
    }

    public void addStockItem(int partIndex, StockItem stockItem) {
        if (po == null)
            po = stockItem.getPo();

        //noinspection unchecked
        List<StockItem> stockItemList = (List<StockItem>) kitStockItemList[partIndex];
        if (stockItemList == null) {
            stockItemList = new ArrayList<>();
            kitStockItemList[partIndex] = stockItemList;

            if (stockItem.getPurchaseItemList().size() > 0)
                matchedPo++;
        }

        stockItemList.add(stockItem);

        dc = Math.min(dc, stockItem.getDc());
    }

    public void addApplySet(int set, String po, int[] ratio) {
        applySet += set;

        for (int i = 0; i < ratio.length; i++) {
            int expectQty = set * ratio[i];
            //noinspection unchecked
            List<StockItem> stockItemList = (List<StockItem>) kitStockItemList[i];
            for (StockItem stockItem: stockItemList) {
                if (expectQty <= 0)
                    break;

                int stockQty = stockItem.getStockQty();
                int oldApplyQty = stockItem.getApplyQty();
                int applyQty = Math.min(expectQty, stockQty - oldApplyQty);

                undoManager.appendInput(stockItem.applyQtyProperty(), applyQty + oldApplyQty);

                String remark = stockItem.getRemark();
                if (remark.isEmpty())
                    remark += REMARK_HEAD + po;
                else
                    remark += ", " + po;

                undoManager.appendInput(stockItem.remarkProperty(), remark);

                expectQty -= applyQty;
            }
        }
    }

    public void clearRedundantRemark() {
        for (int i = 0; i < kitStockItemList.length; i++) {
            List<StockItem> stockItemList = getKitStockItemList(i);
            if (stockItemList == null)
                continue;

            for (StockItem stockItem: stockItemList) {
                if (stockItem.getRemark().equals(REMARK_HEAD + stockItem.getPo())) {
                    undoManager.appendInput(stockItem.remarkProperty(), "");
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
            List<StockItem> stockItemList = (List<StockItem>) kitStockItemList[i];
            if (stockItemList == null) {
                maxSet = 0;
                break;
            }

            int totalStockQty = 0;
            for (StockItem stockItem : stockItemList) {
                totalStockQty += stockItem.getStockQty();
            }
            maxSet = Math.min(maxSet, totalStockQty / ratio[i]);
        }
    }

    public String getPo() {
        return po;
    }

    public int getMaxSet() {
        return maxSet;
    }

    public int getTargetSet() {
        return targetSet;
    }

    public void setTargetSet(int targetSet) {
        this.targetSet = targetSet;
    }

    public int getApplySet() {
        return applySet;
    }

    public void setApplySet(int applySet) {
        this.applySet = applySet;
    }

    public int getMatchedPo() {
        return matchedPo;
    }

    public int getDc() {
        return dc;
    }

    public int getRemainderSet() {
        return targetSet - applySet;
    }

    public List<StockItem> getKitStockItemList(int partIndex) {
        //noinspection unchecked
        return (List<StockItem>) kitStockItemList[partIndex];
    }

    public static List<StockPoNode> getPoNodeList(RotateItem[] kit, int[] ratio, int dcMinLimit) {
        assert kit.length == ratio.length;

        THashMap<String, StockPoNode> poNodeMap = new THashMap<>();

        for (int i = 0; i < kit.length; i++) {
            for (StockItem stockItem : kit[i].getStockItemObsList()) {
                if (stockItem.getDc() < dcMinLimit) {
                    continue;
                }

                String key = stockItem.getPo();
                StockPoNode poNode = poNodeMap.get(key);
                if (poNode == null) {
                    poNode = new StockPoNode(kit.length);
                    poNodeMap.put(key, poNode);
                }

                poNode.addStockItem(i, stockItem);
            }
        }

        List<StockPoNode> poNodeList = new ArrayList<>(poNodeMap.values());
        for (StockPoNode poNode : poNodeList) {
            poNode.postProcess(ratio);
        }

        return poNodeList;
    }

}
