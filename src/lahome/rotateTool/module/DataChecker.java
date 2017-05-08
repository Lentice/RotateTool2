package lahome.rotateTool.module;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.util.Map;

public class DataChecker {
    private static final Logger log = LogManager.getLogger(DataChecker.class.getName());

    public static int doCheck(RotateCollection collection, StringBuilder errorLog) {
        int retValue = 0;
        retValue += checkRatio(collection, errorLog);

        return retValue;
    }

    private static int checkRatio(RotateCollection collection, StringBuilder errorLog) {
        int retValue = 0;

        for (Object o : collection.getKitNodeMap().entrySet()) {
            Map.Entry entry = (Map.Entry) o;
            KitNode kitNode = (KitNode) entry.getValue();

            int expectSet = 0;
            for (RotateItem rotateItem : kitNode.getPartsList()) {
                int ratio = rotateItem.getRatio();
                if (ratio <= 0) {
                    errorLog.append(String.format("ERROR: Ratio is %d \t(Kit: %s, Part: %s)\n",
                            ratio, rotateItem.getKitName(), rotateItem.getPartNumber()));
                    retValue++;
                    rotateItem.setRotateValid(false);
                    continue;
                }

                int pmQty = rotateItem.getPmQty();
                if ((pmQty % ratio) > 0) {
                    errorLog.append(String.format("ERROR: PM Qty (%d) is not multiple of ratio (%d)\t(Kit: %s, Part: %s)\n",
                            pmQty, ratio, rotateItem.getKitName(), rotateItem.getPartNumber()));
                    retValue++;
                    rotateItem.setRotateValid(false);
                    continue;
                }

                int set = pmQty / ratio;
                if (expectSet > 0 && set != expectSet) {
                    errorLog.append(String.format("ERROR: Ratio should be %d \t(Kit: %s, Part: %s)\n",
                            expectSet, rotateItem.getKitName(), rotateItem.getPartNumber()));
                    retValue++;
                    rotateItem.setRotateValid(false);
                    continue;
                }
                expectSet = set;
            }
        }

        return retValue;
    }
}
