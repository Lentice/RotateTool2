package lahome.rotateTool.Util;

import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * Created by Administrator on 2017/5/14.
 */
public class CalculateRotate {

    static int dcMin;

    public static void doCalculate(RotateCollection collection, int dcMinLimit) {
        dcMin = dcMinLimit;

        for (RotateItem rotateItem : collection.getRotateObsList()) {
            if (rotateItem.isDuplicate())
                continue;

            if (!rotateItem.isRotateValid())
                continue;

            RotateItem[] kit = getKitItems(rotateItem);
            if (kit == null)
                continue;

            processStockItems(kit);
        }
    }

    private static void processStockItems(RotateItem[] kit) {
        int[] pmQty = getPmQty(kit);
        int[] ratio = getRatio(kit);

        if (ratio == null)
            return;

        List<StockItem> stockList = kit[0].getStockItemObsList();
        stockList.sort((o1, o2) -> o2.getDc() - o1.getDc());

        for (StockItem stockItem : stockList) {
            if (stockItem.getDc() < dcMin)
                continue;

            String po = stockItem.getPo();
//            stockSet = getSetFromStock(po, kit);
//            purchaseSet = getSetFromPurchase(po, kit);

        }
    }

    private static int[] getRatio(RotateItem[] kit) {
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


    private static int[] getPmQty(RotateItem[] kit) {
        int[] pmQty = new int[kit.length];
        for (int i = 0; i < kit.length; i++) {
            pmQty[i] = kit[i].getPmQty();
        }
        return pmQty;
    }

    public static RotateItem[] getKitItems(RotateItem rotateItem) {
        if (rotateItem.isKit()) {
            if (rotateItem.isFirstPartOfKit())
                return (RotateItem[]) rotateItem.getKitNode().getPartsList().toArray();
            else
                return null;
        } else {
            RotateItem[] kit = new RotateItem[1];
            kit[0] = rotateItem;
            return kit;
        }
    }
}
