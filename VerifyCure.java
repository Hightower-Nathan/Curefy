package CurefyPkg;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** 
* @Author Name: Nathan Hightower
* @Project Name: Curefy
* @Date: Feb 9, 2025
* @Subclass VerifyCure Description: This will handle all of the verification
* steps 
*/
//Imports
//Begin Subclass VerifyCure
public class VerifyCure extends ExcelHandler{
    
   public void complianceFirstHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startFHoldIndex,
            int endFirstHoldIndex) {

        List<Double> lowTcs = new ArrayList<>(); // To store tcs that did not make temp
        List<Double> highTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>(); 
        int dashNumber = 0;
        int startPoint = startFHoldIndex; 
        //Open workbook/sheet
       // try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
         //   Sheet sheet = workbook.getSheetAt(0);
         Sheet sheet = openExcel(dataFilePath);
            System.out.println(startFHoldIndex);
            System.out.println(endFirstHoldIndex);
            
            //System.out.println(rowIndex);
            //Loop through the columns between the start/end indexs and 
            //compare the temps to the spec requirements
            do {
                Row row = sheet.getRow(startFHoldIndex);
                if (row != null) {
                    for (int item : columnIndex) {
                        Cell cell = row.getCell(item);

                        if (dashNumber == 21) {
                            dashNumber = 1;
                        } else {
                            dashNumber++;
                        }
                        if (cell.getNumericCellValue() <= 114.9) {
                            lowTcs.add(cell.getNumericCellValue());
                            failedLowTcNames.add(tcNames.get(dashNumber - 1)); 
                            System.out.printf("\nRow: %d-%d: Low TC found: TC: %s: Temp: %.1f", startFHoldIndex, dashNumber, tcNames.get(dashNumber-1),cell.getNumericCellValue());
                        } else if (cell.getNumericCellValue() >= 145.1) {
                            highTcs.add(cell.getNumericCellValue());
                            failedHighTcNames.add(tcNames.get(dashNumber - 1)); 
                            System.out.printf("\nRow: %d-%d: High TC found: TC: %s: Temp: %.1f", startFHoldIndex, dashNumber, tcNames.get(dashNumber-1),cell.getNumericCellValue());
                        } else {
                            System.out.printf("\nRow: %d-%d Tc's passed requirements...", startFHoldIndex, dashNumber);
                        }
                    }

                }
                startFHoldIndex++;
            } while (startFHoldIndex != endFirstHoldIndex);// outter for loop 

        //} catch (IOException e) {
          //  e.printStackTrace();
       // }
        // findendFirstHold(curespec, dataFilePath, startFHoldIndex, columnIndex, rowIndex);
        //findDelta(curespec, rowIndex, columnIndex, dataFilePath, startPoint,
           // endFirstHoldIndex);
        //writeToReport(lowTcs, highTcs, failedLowTcNames,failedHighTcNames, dataFilePath);
    };
   
   
   
   
    //public void verifyFirstHold(){};
    //public void identifySecondHold(){};
    //public void verifySecondHold(){};
    //public void identifyFinalHold(){};
    //public void verifyFinalHold(){};
    //public void verifyRampRate(){};
    //public void verifyDelta(){};
    //public void verify9002(){};
    //public void verifyVacuum(){}; 
    //public void verifyCoolDown(){};
    
    
    
    
    
    
    
    
    
    
    
    
} //End Subclass VerifyCure
