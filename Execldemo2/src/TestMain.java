import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author Wen xiao
 * @time 2019/10/22
 */

public class TestMain {

    public static void main(String[] args)
    {


        try {
            String url = "D:/ExcelExamRead5.xlsx";

            File aFile = new File(url);

            FileInputStream fio = new FileInputStream(url);

            //创建一个工作薄，实用inputstream创建
            XSSFWorkbook sXssfWorkbook = new XSSFWorkbook(fio);
            XSSFSheet xssfSheet = sXssfWorkbook.getSheetAt(0);


            String value = xssfSheet.getRow(0).getCell(0).getRawValue();

            //这就是改写单元格的方法
            xssfSheet.getRow(2).getCell(2).setCellValue(value+"aaaaaaaa");

            //在outputstream之前将inputstream关闭
            //fio.close();

            //创建outputstream
            FileOutputStream fileOutputStream = new FileOutputStream(aFile);

            //向该工作薄中写入
            sXssfWorkbook.write(fileOutputStream);
            sXssfWorkbook.close();
            fileOutputStream.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}
