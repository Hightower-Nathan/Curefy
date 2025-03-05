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

//Begin Class Curefy
public class Curefy {
        
//Begin Main Method
public static void main(String[] args) { 

   //CureSpec/DataFile paths
   String excelFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/Cure_Specs.xlsx";
   String dataFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/Test_Data_Curefy.xlsx";
   int columnIndex = 1;

   //Create object
   ExcelHandler excelhandler = new ExcelHandler();
   //ColumnValues to returned from readColumn method
   List<Double> columnValues = excelhandler.readColumn(excelFilePath, columnIndex);
   //Give columnValues to the curespec to initialize
   CureSpec curespec = new CureSpec(columnValues);
   //Give ex1 the curespec
   ExcelHandler ex1 = new ExcelHandler(curespec); 
  
   //Find the row index (where to begin)
   int returnedRowIndex = excelhandler.findRowIndex(dataFilePath);
   //Assign TC names
   List<String> tcNames = excelhandler.readTcNames(dataFilePath, returnedRowIndex);
   //Find column index (where to look) 
   List<Integer> returnedColumnIndex = excelhandler.findColumnIndex(dataFilePath, returnedRowIndex);
   //Perform first review 
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
