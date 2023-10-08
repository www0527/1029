import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.List;

public class ExcelUtil {

    public static void readFile(String fileName) throws IOException{

//        讀入檔案，建立excel檔案
        InputStream is = new FileInputStream(new File(fileName));
        XSSFWorkbook wb = new XSSFWorkbook(is);
//        讀取第一個sheet
        XSSFSheet sheet = wb.getSheetAt(0);
//        回傳一個Iterator
        Iterator itRow = sheet.rowIterator();

//        使用Iterator讀取每一個Row
//        .hasNext 若有下一行會繼續，無則跳出
        while(itRow.hasNext()){
//           .next回傳裡面的值，但因row為Object，所要轉型
            XSSFRow row = (XSSFRow)itRow.next();
//            每一行有很多資料儲存格，在乎叫一次取得Iterator
            Iterator itCell = row.cellIterator();

            while(itCell.hasNext()){
                XSSFCell cell = (XSSFCell)itCell.next();
                System.out.print(cell + " ");
            }
            System.out.println("");
        }
    }
    public static void writeFile(String fileName, List<String[]> dataList) throws IOException{
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("Test sheet");
        XSSFRow titleRow = sheet.createRow(0);

//        製作title欄位
        XSSFCell cell1 = titleRow.createCell(0);
        cell1.setCellValue("id");
        XSSFCell cell2 = titleRow.createCell(1);
        cell2.setCellValue("name");
        XSSFCell cell3 = titleRow.createCell(2);
        cell3.setCellValue("age");

//        寫入EXcel
        for(int i= 0; i < dataList.size(); i++){
            XSSFRow row = sheet.createRow(i+1);
            String[] rowData = dataList.get(i);

            for(int j=0; j<rowData.length; j++){
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(rowData[j]);
            }
        }

        FileOutputStream fis = new FileOutputStream(fileName);
        wb.write(fis);
        fis.flush();
        fis.close();
    }
}
