import java.util.List;
import java.util.Map;

public class StartUtil {

    public static List<List<Double>> getSheBanData(String xlsPath) throws Exception {
        return XlsxUtil.get2ColData(xlsPath, 0, 0);
    }

    public static Map<List<List<Double>>, List<List<Double>>> getGuDingDaoYeData(String xlsPath) throws Exception {
        int guDingStartRow = 2;
        int guDingStartColHead = 4;
        int guDingStartColTail = 6;
        return XlsxUtil.getGuDingDaoYeData(xlsPath, guDingStartRow, guDingStartColHead, guDingStartColTail);
    }

    public static Map<Integer, List<List<Double>>> getPatternArgs(String xlsPath) throws Exception {
        int patternArgStartRow = 1;
        int patternArgStartCol = 8;
        return XlsxUtil.getPatternArgs(xlsPath, patternArgStartRow, patternArgStartCol);
    }
}
