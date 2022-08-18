import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class XlsxUtil {

    private static XSSFSheet createSheet(String xlsxPath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(xlsxPath);
        return new XSSFWorkbook(fileInputStream).getSheetAt(0);
    }

    public static List<List<Double>> get2ColData(String xlsxPath, int startRow, int startCol) throws IOException {
        XSSFSheet sheet = createSheet(xlsxPath);
        List<List<Double>> result = new ArrayList<>();
        for (int rowIndex = startRow; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            List<Double> inner = new ArrayList<>();
            //两列数据
            for (int colIndex = startCol; colIndex < 2 + startCol; colIndex++) {
                Cell rowCell = sheet.getRow(rowIndex).getCell(colIndex);
                //非空校验
                if (rowCell == null) break;
                //类型校验
                if (rowCell.getCellTypeEnum() != CellType.NUMERIC) continue;
                inner.add(rowCell.getNumericCellValue());
            }
            //所有的列数据行数不相同
            if (inner.size() != 0) result.add(inner);
        }
        return result;
    }

    public static Map<List<List<Double>>,List<List<Double>>> getGuDingDaoYeData(String xlsxPath) throws Exception {
        //get sheet
        XSSFSheet sheet = createSheet(xlsxPath);
        //row 1
        XSSFRow rowTitle = sheet.getRow(0);
        int kindCount = 0;
        //get kind count
        for (int i = 0; i < rowTitle.getLastCellNum(); i++) {
            Cell cell = rowTitle.getCell(i);
            if (cell==null) continue;
            if (cell.getCellTypeEnum() == CellType.STRING) {
                if ("尾部连接".equals(cell.getStringCellValue())) kindCount++;
            }
        }
        //固定导叶数据
        Map<List<List<Double>>,List<List<Double>>> guDingData=new HashMap<>();
        int guDingStartRow=2;
        int guDingStartColHead=4;
        int guDingStartColRow=6;
        int evertCountCol=7;
        for (int colN = 0; colN < kindCount; colN++) {
            List<List<Double>> headData = get2ColData(xlsxPath, guDingStartRow, guDingStartColHead+ (evertCountCol* colN));
            List<List<Double>> tailData = get2ColData(xlsxPath, guDingStartRow, guDingStartColRow + (evertCountCol * colN));
            guDingData.put(headData,tailData);
        }
        return guDingData;
    }
}
