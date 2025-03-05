package CurefyPkg;
/**
* @Author Name: Nathan Hightower
* @Project Name: Curefy
* @Date: Feb 8, 2025
* @Subclass excelHandler Description: This will handle all of the excel stuff
*/
//Imports
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException; 
import java.util.ArrayList;
import java.util.List;

import java.awt.Desktop;
import java.io.File;

//Begin Subclass ExcelHandler
public class ExcelHandler extends CureSpec{
    
    //Use method chainning for all of these 
    public ExcelHandler(){};
    public ExcelHandler(CureSpec curespec){};
    
    int startFirstHoldIndex, startSecondHoldIndex, startThirdHoldIndex; 
    int endFirstHoldIndex, endSecondHoldIndex, endThirdHoldIndex;
    List<String> tcNames = new ArrayList<>(); 
    
    
    
    //This find the correct rowIndex 
    public int findRowIndex (String dataFilePath){
        int rowIndex = 0;
        try (FileInputStream fis2 = new FileInputStream(dataFilePath);
               Workbook workbook = new XSSFWorkbook(fis2)){
                   Sheet sheet = workbook.getSheetAt(0);
                   for(Row row: sheet){
                       for(Cell cell: row){
                           if(cell.getCellType() == CellType.STRING && 
                                   cell.getStringCellValue().contains("Time"))
                                  rowIndex = row.getRowNum();
                                  break;
                       }   
                   }
        }catch (IOException e){
                       e.printStackTrace();
                       }
        //System.out.println(rowIndex);
        return rowIndex; 
    };
    
    public List<Integer> findColumnIndex (String dataFilePath, int rowIndex){
        int columnIndexx = 0;
        List<Integer> columnIndex = new ArrayList<>();
        try (FileInputStream fis2 = new FileInputStream(dataFilePath);
               Workbook workbook = new XSSFWorkbook(fis2)){
                   Sheet sheet = workbook.getSheetAt(0);
                   Row row = sheet.getRow(rowIndex);
                   if (row != null){
                   //for(Row row: sheet){
                       for(Cell cell: row){
                           if(cell.getCellType() == CellType.STRING && 
                                   cell.getStringCellValue().contains("PTC")){
                           columnIndexx = cell.getColumnIndex();
                           columnIndex.add(columnIndexx);
                           }
                                   
                       }
                   }  
        }catch (IOException e){
                       e.printStackTrace();
                       }
        //for(int item: columnIndex){
         //   System.out.println(item);
        //}
       
        return columnIndex; 
    };
   
    //Works and locates the correct row!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    public List<String> readTcNames (String dataFilePath, int rowIndex){
        
       // List<String> tcNames = new ArrayList<>();
        try (FileInputStream fis2 = new FileInputStream(dataFilePath);
               Workbook workbook = new XSSFWorkbook(fis2)){
                   Sheet sheet = workbook.getSheetAt(0);
                   Row row = sheet.getRow(rowIndex);
                   if (row != null){
                   //for(Row row: sheet){
                       for(Cell cell: row){
                           if(cell.getCellType() == CellType.STRING && 
                                   cell.getStringCellValue().contains("PTC")){
                                   tcNames.add(cell.getStringCellValue()); 
                           }
                                   
                       }
                   }  
        }catch (IOException e){
                       e.printStackTrace();
                       }
       // tcNames = tcNames;
        return tcNames; 
    };
    
    /**
     * Method: Reads the data on column 1 of the specified cure spec
     * @param excelFilePath
     * @param columnIndex
     * @return 
     */
    public List<Double> readColumn(String excelFilePath, int columnIndex){
       List<Double> columnData = new ArrayList<>();//<String>
       try (FileInputStream fis = new FileInputStream(excelFilePath);
               Workbook workbook = new XSSFWorkbook(fis)){
                   Sheet sheet = workbook.getSheetAt(0);
                   for(Row row: sheet){
                       Cell cell = row.getCell(columnIndex);
                       if(cell != null){
                           columnData.add(cell.getNumericCellValue());//cell.toString()
                       }
                   }
               } catch (IOException e){
                       e.printStackTrace();
                       }
       
               return columnData;
   }//End method 
    
    
    
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    //This grouping works - method chain the rest of them
    //This can be established by using two methods if planned correctly
    //instead of six. Do this at some point for now, move on. 
    
    
   public void findFirstHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath) {
        int startFHoldIndex = 0;
        int startRowInd = rowIndex + 2;
        double condition = 0.0;

        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
            do {
                Row row = sheet.getRow(startRowInd);
                if (row != null) {
                    for (int item : columnIndex) {
                        Cell cell = row.getCell(item);
                        if (cell.getNumericCellValue() <= 115) {
                            condition = cell.getNumericCellValue();
                            if (condition >= 115) {
                                startFHoldIndex = row.getRowNum();
                                startFirstHoldIndex = startFHoldIndex;
                               
                            }

                        }
                    }
                }
                startRowInd++;
            } while (condition <= 114.9);// outter for loop 
        } catch (IOException e) {
            e.printStackTrace();
        }
            findendFirstHold(curespec, dataFilePath, startFHoldIndex, columnIndex, rowIndex);
        
    };
    
   public void findendFirstHold(CureSpec curespec,String datafilePath, int startFirstHoldIndex, List<Integer> columnIndex, int rowIndex ){
    int endFHoldIndex = 0; 
    double bHoldTimeMinutes = curespec.getbHoldTime(); 
    double elapsedMinutes = 0.0;
    int originalValue = startFirstHoldIndex + 1;
    
     try (FileInputStream fis = new FileInputStream(datafilePath);
               Workbook workbook = new XSSFWorkbook(fis)){
                   Sheet sheet = workbook.getSheetAt(0);
                   do{
                       Row row = sheet.getRow(startFirstHoldIndex);
                       
                       //Just looks at the column with the number of minutes 
                       for(int item : columnIndex){
                           Cell cell = row.getCell(item);           
                       if(cell != null){
                           endFHoldIndex = row.getRowNum();
                       }
                       
                        if (elapsedMinutes == bHoldTimeMinutes){  
                               System.out.println("***************************");
                               System.out.println("**First Hold Identified");
                               System.out.printf("**Start Hold Index: %d\n", originalValue);
                               System.out.printf("**End of Hold Index: %d\n", endFHoldIndex);
                               System.out.printf("**Number of minutes: %.1f\n",elapsedMinutes);
                               System.out.println("***************************");
                              
                               break;
                           } 
                       }
                       elapsedMinutes++;startFirstHoldIndex++;
                   }while(elapsedMinutes != bHoldTimeMinutes + 1); 
                       } catch (IOException e){
                       e.printStackTrace();
               }
     
     endFirstHoldIndex = endFHoldIndex; 
     complianceFirstHold(curespec, rowIndex, columnIndex, datafilePath, originalValue,//was startFirstHoldIndex
             endFirstHoldIndex);
     
     //Commenting out just to test the first hold reqs 
     //findSecondHold(rowIndex, columnIndex, datafilePath, curespec);
 
    
    };//End method 
     
   //WIP compare the temps for compliance during the first hold
   //Fucking works..Hell yeah brother 
   public void complianceFirstHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startFHoldIndex,
            int endFirstHoldIndex) {

        List<Double> lowTcs = new ArrayList<>(); // To store tcs that did not make temp
        List<Double> highTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>(); 
        int dashNumber = 0;
        
        //Open workbook/sheet
        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
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

        } catch (IOException e) {
            e.printStackTrace();
        }
        // findendFirstHold(curespec, dataFilePath, startFHoldIndex, columnIndex, rowIndex);
        
        writeToReport(lowTcs, highTcs, failedLowTcNames,failedHighTcNames, dataFilePath);
    };
   
   
   
   
   public void writeToReport (List<Double> lowTcs, List<Double> highTcs, List<String> failedLowTcNames, List<String> failedHighTcNames, String dataFilePath){
   
    
   int increaseRow = 0;
   int increaseColumn = 0; 
  
   
   try (FileInputStream fis = new FileInputStream(dataFilePath);  
           Workbook workbook = new XSSFWorkbook(fis)) {
       
            Sheet sheet = workbook.createSheet("Curefy_Report");
            
            //Make a loop to increase the row and column num
            Row row = sheet.createRow(increaseRow); //increase row
            Cell cell1 = row.createCell(increaseColumn); // increase cell in row
            cell1.setCellValue("First Hold");

            for(int i = 0; i < failedLowTcNames.size(); i++){
                increaseRow = increaseRow+1;
                row = sheet.createRow(increaseRow);
                cell1 = row.createCell(increaseColumn);//added to test 
                cell1.setCellValue(failedLowTcNames.get(i));
                cell1 = row.createCell(increaseColumn+1);
                cell1.setCellValue(lowTcs.get(i));
               
            
            };
            
           
            try (FileOutputStream fos = new FileOutputStream(dataFilePath)){
                workbook.write(fos);
                workbook.close();
            }
       Desktop.getDesktop().open(new File(dataFilePath));//View the report 
       
   }catch (IOException e) {
            e.printStackTrace();
        }
  
   //Desktop.getDesktop().open(new File(dataFilePath));
   }
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
     
     //Previous return method keep for now - but use the void method 
   
   /*public int dfindendFirstHold(CureSpec curespec,String datafilePath, int startFirstHoldIndex, List<Integer> columnIndex ){
    int endFHoldIndex = 0; 
    double bHoldTimeMinutes = curespec.getbHoldTime(); 
    double elapsedMinutes = 0.0;
    int originalValue = startFirstHoldIndex + 1;
     try (FileInputStream fis = new FileInputStream(datafilePath);
               Workbook workbook = new XSSFWorkbook(fis)){
                   Sheet sheet = workbook.getSheetAt(0);
                   do{
                       Row row = sheet.getRow(startFirstHoldIndex);
                       Cell cell = row.getCell(columnIndex);                   
                       if(cell != null){
                           endFHoldIndex = row.getRowNum();
                       }startFirstHoldIndex++;elapsedMinutes++;
                        if (elapsedMinutes == bHoldTimeMinutes){  
                               System.out.println("***************************");
                               System.out.println("***First Hold Identified***");
                               System.out.printf("Start Hold Index: %d\n", originalValue);
                               System.out.printf("End of Hold Index: %d\n", endFHoldIndex + 1);
                               System.out.printf("Number of minutes: %.1f\n",elapsedMinutes);
                               System.out.println("***************************");
                              
                               break;
                           }                      
                   }while(bHoldTimeMinutes != elapsedMinutes);
                       } catch (IOException e){
                       e.printStackTrace();
               }
     
 
    return endFHoldIndex; 
    };//End method 
    
    
    
   public void findSecondHold(int rowIndex, List<Integer> columnIndex, String dataFilePath, CureSpec curespec) {
        int startSHoldIndex = 0;
        int startRowInd = rowIndex + 197;
        double condition = 0.0;

        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
            do {
                Row row = sheet.getRow(startRowInd);
                if (row != null) {
                    for (int item : columnIndex) {
                        Cell cell = row.getCell(item);
                        
                        if (cell.getNumericCellValue() <= 240) {
                            condition = cell.getNumericCellValue();
                            if (condition >= 240) {
                                startSHoldIndex = row.getRowNum();
                                startSecondHoldIndex = startSHoldIndex;
                            }

                        }
                    }
                }
                startRowInd++;
            } while (condition <= 239.9);// outter for loop 
        } catch (IOException e) {
            e.printStackTrace();
        }

        findendSecondHold(curespec, dataFilePath, startSHoldIndex, columnIndex, rowIndex); 
        
    };
    
    
   public void findendSecondHold(CureSpec curespec,String datafilePath, int startSecondHoldIndex, List<Integer> columnIndex, int rowIndex){
    int endSHoldIndex = 0; 
    double dHoldTimeMinutes = curespec.getdHoldTime(); 
    double elapsedMinutes = 0.0;
    int originalValue = startSecondHoldIndex + 1;
    
     try (FileInputStream fis = new FileInputStream(datafilePath);
               Workbook workbook = new XSSFWorkbook(fis)){
                   Sheet sheet = workbook.getSheetAt(0);
                   do{
                       Row row = sheet.getRow(startSecondHoldIndex);
                       for(int item : columnIndex){
                       Cell cell = row.getCell(item);   
                       if(cell != null){
                           endSHoldIndex = row.getRowNum();
                       }
                       
                        if (elapsedMinutes == dHoldTimeMinutes){  
                               System.out.println("***************************");
                               System.out.println("***Second Hold Identified***");
                               System.out.printf("Start Hold Index: %d\n", originalValue);
                               System.out.printf("End of Hold Index: %d\n", endSHoldIndex);
                               System.out.printf("Number of minutes: %.1f\n",elapsedMinutes);
                               System.out.println("***************************");
                               break;
                           } 
                       }
                       elapsedMinutes++;startSecondHoldIndex++;
                   }while(elapsedMinutes != dHoldTimeMinutes + 1 );
                       } catch (IOException e){
                       e.printStackTrace();
               }
     
    endSecondHoldIndex = endSHoldIndex;
    findThirdHold(rowIndex, columnIndex, datafilePath, curespec);
    };
    
    
   public void findThirdHold(int rowIndex, List<Integer> columnIndex, String dataFilePath, CureSpec curespec) {
        int startTHoldIndex = 0;
        int startRowInd = rowIndex + 455; // can use the previous end hold index
        double condition = 0.0;

        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
            do {
                Row row = sheet.getRow(startRowInd);
                if (row != null) {
                    for (int item : columnIndex) {
                        Cell cell = row.getCell(item);
                        
                        if (cell.getNumericCellValue() <= 290) {
                            condition = cell.getNumericCellValue();
                            if (condition >= 290) {
                                startTHoldIndex = row.getRowNum();
                                startThirdHoldIndex = startTHoldIndex; 
                            }

                        }
                    }
                }
                startRowInd++;
            } while (condition <= 289.9);// outter for loop 
        } catch (IOException e) {
            e.printStackTrace();
        }

        findendThirdHold(curespec, dataFilePath, startTHoldIndex, columnIndex, rowIndex);
    };
    
   public void findendThirdHold(CureSpec curespec,String datafilePath, int startThirdHoldIndex, List<Integer> columnIndex, int rowIndex){
    int endTHoldIndex = 0; 
    double fHoldTimeMinutes = curespec.getfHoldTime(); 
    double elapsedMinutes = 0.0;
    int originalValue = startThirdHoldIndex + 1;
    
     try (FileInputStream fis = new FileInputStream(datafilePath);
               Workbook workbook = new XSSFWorkbook(fis)){
                   Sheet sheet = workbook.getSheetAt(0);
                   do{
                       Row row = sheet.getRow(startThirdHoldIndex);
                       for(int item : columnIndex){
                       Cell cell = row.getCell(item);                   
                       if(cell != null){
                           endTHoldIndex = row.getRowNum();
                       }
                       
                        if (elapsedMinutes == fHoldTimeMinutes){  
                               System.out.println("***************************");
                               System.out.println("***Third Hold Identified***");
                               System.out.printf("Start Hold Index: %d\n", originalValue);
                               System.out.printf("End of Hold Index: %d\n", endTHoldIndex);
                               System.out.printf("Number of minutes: %.1f\n",elapsedMinutes);
                               System.out.println("***************************");
                               
                               break;
                           } 
                       }
                       elapsedMinutes++;startThirdHoldIndex++; 
                   }while(elapsedMinutes != fHoldTimeMinutes + 1);
                       } catch (IOException e){
                       e.printStackTrace();
               }
     

    
    };
    
    
    
    
    
    
    
    
    
    
    /*Method for creating excel sheet for report generation here*/
    
    /*Method for writing to the compliance report*/
    
    /*Method for creating 9002 report here...if already created, open*/
    
    /*Method for writing to the 9002 report*/
    
    /*Method for opening and displaying the reports*/
    
    
    
    
    
    
    
    
   
  
    
    
} //End Subclass ExcelHandler
