import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

/**
 * @author Wen xiao
 * @time 2019/10/22
 */
public class Caculat {


    /**
     * @param args
     */


    public static void main(String[] args) throws Exception {


        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("D:/ExcelExamRead5.xlsx"));
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rsRows = sheet.getPhysicalNumberOfRows();// 行
            XSSFRow hssfRow = sheet.getRow(0);//得到一行，进而下一步得到列数
            int rsColumns = hssfRow.getPhysicalNumberOfCells();// 列
            for (int i = 0; i < rsRows; i++)
            {
                for (int j = 0; j < rsColumns; j++)
                {
                    XSSFCell c1 = sheet.getRow(i).getCell(j);// 获取第i行数据的第j列
                    double p = c1.getNumericCellValue();
                    System.out.print(p+",");
                }
                System.out.println();
            }
        } catch (Exception e) {


            throw new RuntimeException(e);


        }


    }


}

