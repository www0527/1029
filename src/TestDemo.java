import java.io.IOException;
import java.util.List;

public class TestDemo {

    public static void main(String[] args) throws IOException {

////        讀取Excel資料
//        String readFileName ="C:\Users\willy\IdeaProjects\Java13_3\read.xlsx";
//        try {
//            ExcelUtil.readFile(readFileName);
//        }catch (IOException e){
//            e.printStackTrace();
//        }

//        寫入Excel資料
        String writeFileName = "C:\\Users\\willy\\IdeaProjects\\Java13_3\\write.xlsx";
        DataHelp dataHelp = new DataHelpImp();
        List<String[]> list = dataHelp.getData();
        try {
            ExcelUtil.writeFile(writeFileName, list);
        }catch (IOException e){
            e.printStackTrace();
        }
        System.out.println("Write Excel End");
    }
}
