package CurefyPkg;
/** 
* @Author Name: Nathan Hightower
* @Project Name: Curefy
* @Date: Jan 25, 2025
* @Description: This is for testing the start of my Capstone project 
*/
//Imports

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException; 
import java.util.ArrayList;
import java.util.List;

//import java.util.concurrent.Executors; for thread manager since too many threads can downgrade preformance


//Begin Class Curefy
public class Curefy {
        
//Begin Main Method
public static void main(String[] args) { //Use multithreading like this assign each task to a thread 
   /* int n = 8;
    for(int i = 0; i < n; i++){
        MyRunnable myrun = new MyRunnable() {};
        Thread thread = new Thread(myrun);
        thread.start();
    }*/
  
   String excelFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/Cure_Specs.xlsx";
   String dataFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/Test_Data_Curefy(1).xlsx";

   int columnIndex = 1;
   
   ExcelHandler excelhandler = new ExcelHandler();
   List<Double> columnValues = excelhandler.readColumn(excelFilePath, columnIndex);
   CureSpec curespec = new CureSpec(columnValues);
   ExcelHandler ex1 = new ExcelHandler(curespec); 
   
  
   
   int returnedRowIndex = excelhandler.findRowIndex(dataFilePath);
   List<String> tcNames = excelhandler.readTcNames(dataFilePath, returnedRowIndex);
   List<Integer> returnedColumnIndex = excelhandler.findColumnIndex(dataFilePath, returnedRowIndex);
   
   excelhandler.findFirstHold(curespec, returnedRowIndex, returnedColumnIndex, dataFilePath);
   
   //int firstHoldIndex = excelhandler.findFirstHold(returnedRowIndex, returnedColumnIndex, dataFilePath);
   //System.out.println(firstHoldIndex);
   
   //int firstHoldEndIndex = excelhandler.findendFirstHold(curespec,dataFilePath, firstHoldIndex, columnIndex);
   //System.out.println(firstHoldEndIndex + 1);
  
    //int secondHoldIndex = excelhandler.findSecondHold(returnedRowIndex, returnedColumnIndex, dataFilePath);
    //System.out.println(secondHoldIndex + 1);
    
    //int secondHoldEndIndex = excelhandler.findendSecondHold(curespec,dataFilePath, secondHoldIndex, columnIndex);
   //System.out.println(secondHoldEndIndex + 1);
   
    //int thirdHoldIndex = excelhandler.findThirdHold(returnedRowIndex, returnedColumnIndex, dataFilePath);
   
    //int thirdHoldEndIndex = excelhandler.findendThirdHold(curespec,dataFilePath, thirdHoldIndex, columnIndex);
   
} //End Main Method
} //End Class Curefy
