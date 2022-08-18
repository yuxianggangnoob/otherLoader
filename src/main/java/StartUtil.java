import java.util.List;
import java.util.Map;

public class StartUtil {

    public static List<List<Double>> getSheBanData(String xlsPath) throws Exception {
        return XlsxUtil.get2ColData(xlsPath,0, 0);
    }
    public static Map<List<List<Double>>,List<List<Double>>> getGuDingDaoYeData(String xlsPath) throws Exception {
        return XlsxUtil.getGuDingDaoYeData(xlsPath);
    }
}
