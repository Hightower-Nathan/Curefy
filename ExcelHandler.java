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
import java.util.Iterator;
import java.util.Date;
import java.sql.Time;

//import java.time.LocalDateTime;

//Begin Subclass ExcelHandler
public class ExcelHandler extends CureSpec {

    //Use method chainning for all of these 
    public ExcelHandler() {
    }

    ;
    public ExcelHandler(CureSpec curespec) {
    }
    ;
   
    
    
    int startFirstHoldIndex, startSecondHoldIndex, startThirdHoldIndex;
    int endFirstHoldIndex, endSecondHoldIndex, endThirdHoldIndex;
    List<String> tcNames = new ArrayList<>();
    List<Integer> vacuumIndex = new ArrayList<>();
    List<String> vacNames = new ArrayList<>();
    List<Double> vacInHg = new ArrayList<>();
    int dataStart = 0;
    int dataEnd = 0;
    int dataStart2 = 0;
    private double eMinTemp; 
    private double aMinTemp;
    private double cMinTemp;
    private double aMaxTemp;
    private double cMaxTemp;
    private double eMaxTemp; 
    private double lessTemp;
    private double vacPressure;
    private double minVacPressure; 
    private double temp9002;
    private double rampRateMax; 
    Date dateTime = new Date();
    private String fileName,ovenNum,runRecipe,cureJob; 
    String reportFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/CurefyReportTemplate(7).xlsx";
    

    //This find the correct rowIndex 
    public int findRowIndex(String dataFilePath) {
        int rowIndex = 0;

        
        Sheet sheet = openExcel(dataFilePath);
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Time")) {
                    rowIndex = row.getRowNum();
                    dataStart = rowIndex + 2;// to be used in the vacuum verify portion 
                    dataStart2 = rowIndex + 2;
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Filename:")) {
                    Cell nextTo = row.getCell(1);
                    fileName = nextTo.toString();
                    System.out.println(fileName);
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Equipment:")) {
                    Cell nextTo = row.getCell(1);
                    ovenNum = nextTo.toString();
                    System.out.println(ovenNum);
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Run Recipe:")) {
                    Cell nextTo = row.getCell(1);
                    runRecipe = nextTo.toString();
                    System.out.println(runRecipe);
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Part Number:")) {
                    Cell nextTo = row.getCell(1);
                    cureJob = nextTo.toString();
                    System.out.println(cureJob);
                }
                
  
                break;
            }
        }
        return rowIndex;
    }

    ;
    
    public List<Integer> findColumnIndex(String dataFilePath, int rowIndex) {
        int columnIndexx = 0;
        int vacuumIndexx = 0;
        List<Integer> columnIndex = new ArrayList<>();
        
        Sheet sheet = openExcel(dataFilePath);
        Row row = sheet.getRow(rowIndex);
        if (row != null) {

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("PTC")) {
                    columnIndexx = cell.getColumnIndex();
                    columnIndex.add(columnIndexx);
                } else if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("VPRB")) {
                    vacuumIndexx = cell.getColumnIndex();
                    vacuumIndex.add(vacuumIndexx);
                    vacNames.add(cell.getStringCellValue()); //To store the names of the vacuum TCS for pressure check.  
               }

            }
        }

        return columnIndex;
    }

    ;
   
    //Works and locates the correct row!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    public List<String> readTcNames(String dataFilePath, int rowIndex) {
        
        Sheet sheet = openExcel(dataFilePath);
        Row row = sheet.getRow(rowIndex);
        if (row != null) {

            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("PTC")) {
                    tcNames.add(cell.getStringCellValue());
                }

            }
        }

        return tcNames;
    }

    ;
    
    /**
     * Method: Reads the data on column 1 of the specified cure spec
     * @param excelFilePath
     * @param columnIndex
     * @return 
     */
    public List<Double> readColumn(String excelFilePath, int columnIndex) {
        List<Double> columnData = new ArrayList<>();
        
        Sheet sheet = openExcel(excelFilePath);
        for (Row row : sheet) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null) {
                columnData.add(cell.getNumericCellValue());
            }
        }

        return columnData;
    }//End method 

    
    public void findFirstHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath) {
        int startFHoldIndex = 0;
        int startRowInd = rowIndex + 2;
        double condition = 0.0;
        aMinTemp = curespec.getabMinTemp();
        lessTemp = curespec.getLessTemp(); 
        
        double stopCondition = aMinTemp - lessTemp;  
        

        Sheet sheet = openExcel(dataFilePath);
        do {
            Row row = sheet.getRow(startRowInd);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);
                    if (cell.getNumericCellValue() <= aMinTemp) {

                        condition = cell.getNumericCellValue();
                       
                    
                    if (condition >= aMinTemp) {
                          
                            startFHoldIndex = row.getRowNum();
                            startFirstHoldIndex = startFHoldIndex;
                         
                        }
                    else{break;}
                }

                   
                }
            }
            startRowInd++;
            
        } while (condition <= stopCondition);// outter for loop 
        
        
        startRowInd = rowIndex + 2;
 
        findRampRate(curespec, rowIndex, columnIndex, dataFilePath, startRowInd,
                startFirstHoldIndex, sheet);
       
       findendFirstHold(curespec, dataFilePath, startFHoldIndex, columnIndex, rowIndex);//------------------------------commented to test time

    }; 
    
   public void findendFirstHold(CureSpec curespec, String datafilePath, int startFirstHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endFHoldIndex = 0;
        double bHoldTimeMinutes = curespec.getbHoldTime();
        double elapsedMinutes = 0.0;
        int originalValue = startFirstHoldIndex + 1;
      
        Sheet sheet = openExcel(datafilePath);
        do {
            Row row = sheet.getRow(startFirstHoldIndex);

            //Just looks at the column with the number of minutes 
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endFHoldIndex = row.getRowNum();
                }

                if (elapsedMinutes == bHoldTimeMinutes) {
                    System.out.println("\n***************************");
                   System.out.println("**First Hold Identified");
                   System.out.printf("**Start Hold Index: %d\n", originalValue);
                    System.out.printf("**End of Hold Index: %d\n", endFHoldIndex);
                    System.out.printf("**Number of minutes: %.1f\n", elapsedMinutes);
                    System.out.println("***************************");

                    break;
                }
            }
            elapsedMinutes++;
            startFirstHoldIndex++;
        } while (elapsedMinutes != bHoldTimeMinutes + 1);

        endFirstHoldIndex = endFHoldIndex;

        complianceFirstHold(curespec, rowIndex, columnIndex, datafilePath, originalValue,
                endFirstHoldIndex);

    }

    ;//End method 
     
    
   public void findSecondHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int endFirstHoldIndex) {
        int startSHoldIndex = 0;
        int startRowInd = endFirstHoldIndex + 1;
        double condition = 0.0;
        cMinTemp = curespec.getcdMinTemp();
        
        double stopCondition = cMinTemp - lessTemp;

        Sheet sheet = openExcel(dataFilePath);
        do {
            Row row = sheet.getRow(startRowInd);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);
                    if (cell.getNumericCellValue() <= cMinTemp) {

                        condition = cell.getNumericCellValue();
                        if (condition >= cMinTemp) {
                            startSHoldIndex = row.getRowNum();
                            startSecondHoldIndex = startSHoldIndex;
                        }

                    }
                }
            }
            startRowInd++;
        } while (condition <= stopCondition);// outter for loop 

        findendSecondHold(curespec, dataFilePath, startSHoldIndex, columnIndex, rowIndex);
        

    }

    ;
   
 public void findendSecondHold(CureSpec curespec, String datafilePath, int startSecondHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endSHoldIndex = 0;
        double dHoldTimeMinutes = curespec.getdHoldTime();
        double elapsedMinutes = 0.0;
        int originalValue = startSecondHoldIndex + 1;

        Sheet sheet = openExcel(datafilePath);
        do {
            Row row = sheet.getRow(startSecondHoldIndex);

            //Just looks at the column with the number of minutes 
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endSHoldIndex = row.getRowNum();
                }

                if (elapsedMinutes == dHoldTimeMinutes) {
                    System.out.println("\n***************************");
                   System.out.println("**Second Hold Identified");
                    System.out.printf("**Start Hold Index: %d\n", originalValue);
                   System.out.printf("**End of Hold Index: %d\n", endSHoldIndex);
                   System.out.printf("**Number of minutes: %.1f\n", elapsedMinutes);
                   System.out.println("***************************");

                    break;
                }
            }
            elapsedMinutes++;
            startSecondHoldIndex++;
        } while (elapsedMinutes != dHoldTimeMinutes + 1);

        endSecondHoldIndex = endSHoldIndex;
        complianceSecondHold(curespec, rowIndex, columnIndex, datafilePath, originalValue,//was startFirstHoldIndex
                endSecondHoldIndex);

        //Commenting out just to test the first hold reqs 
         findThirdHold(curespec, rowIndex, columnIndex, datafilePath, endSecondHoldIndex);
    };
   
 public void findThirdHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int endSecondHoldIndex) {
        int startTHoldIndex = 0;
        int startRowInd = endSecondHoldIndex + 1;
        double condition = 0.0;
        eMinTemp = curespec.getefMinTemp();
       
        double stopCondition = eMinTemp - lessTemp;

        Sheet sheet = openExcel(dataFilePath);
        do {
            Row row = sheet.getRow(startRowInd);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);
                    if (cell.getNumericCellValue() <= eMinTemp) {

                        condition = cell.getNumericCellValue();
                        if (condition >= eMinTemp) {
                            startTHoldIndex = row.getRowNum();
                            startThirdHoldIndex = startTHoldIndex;
                        }

                    }
                }
            }
            startRowInd++;
        } while (condition <= stopCondition);// outter for loop 

        findendThirdHold(curespec, dataFilePath, startTHoldIndex, columnIndex, rowIndex);
        

    };
   
 public void findendThirdHold(CureSpec curespec, String datafilePath, int startThirdHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endTHoldIndex = 0;
        double fHoldTimeMinutes = curespec.getfHoldTime();
        double elapsedMinutes = 0.0;
        int originalValue = startThirdHoldIndex + 1;

        Sheet sheet = openExcel(datafilePath);
        do {
            Row row = sheet.getRow(startThirdHoldIndex);

            //Just looks at the column with the number of minutes 
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endTHoldIndex = row.getRowNum();
                }

                if (elapsedMinutes == fHoldTimeMinutes) {
                    System.out.println("\n***************************");
                    System.out.println("**Third Hold Identified");
                    System.out.printf("**Start Hold Index: %d\n", originalValue);
                    System.out.printf("**End of Hold Index: %d\n", endTHoldIndex);
                    System.out.printf("**Number of minutes: %.1f\n", elapsedMinutes);
                    System.out.println("***************************");

                    break;
                }
            }
            elapsedMinutes++;
            startThirdHoldIndex++;
        } while (elapsedMinutes != fHoldTimeMinutes + 1);

        endThirdHoldIndex = endTHoldIndex;
        complianceThirdHold(curespec, rowIndex, columnIndex, datafilePath, originalValue,//was startFirstHoldIndex
                endFirstHoldIndex);

        
    };
   
   //For testing purposes 
 public Sheet openExcel(String dataFilePath) {

        Sheet s = null;
        //System.out.println("\nEntering Open Excel");
        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
            s = sheet;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return s;
    }

    ;
   
//Make this a single function that takes in the indexes and curespec temp ranges 
 public void complianceFirstHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startFHoldIndex,
            int endFirstHoldIndex) {

        List<Double> flowTcs = new ArrayList<>(); // To store tcs that did not make temp
        List<Double> fhighTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        
        List<Time> lowTimeStamp = new ArrayList<>();//to store timestamp
        List<Time> highTimeStamp = new ArrayList<>();//to store timestamp 
        
        int identity = 0;
        int dashNumber = 0;
        int startPoint = startFHoldIndex;
        int locating = 0; 
        aMaxTemp = curespec.getabMaxTemp();
        int j = 0;
     

        Sheet sheet = openExcel(dataFilePath);
       
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
                    if (cell.getNumericCellValue() <= aMinTemp - lessTemp/*114.9*/) {
                        System.out.println("Low temps");
                        
                        flowTcs.add(cell.getNumericCellValue());
                       failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        
                        //was Cell cellB
                        Cell cellB = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                        Time lowTime = new Time(tcTime.getTime());//------------------------------added to test time
                        lowTimeStamp.add(lowTime);//------------------------------added to test time
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", lowTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
  
                    } else if (cell.getNumericCellValue() >= aMaxTemp + lessTemp/*145.1*/) {
                        System.out.println("High temps");
                        
                        fhighTcs.add(cell.getNumericCellValue());
                       failedHighTcNames.add(tcNames.get(dashNumber - 1)); 
                        
                       Cell cellC = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellC.getDateCellValue();//------------------------------added to test time
                        Time highTime = new Time(tcTime.getTime());//------------------------------added to test time
                        highTimeStamp.add(highTime);//------------------------------added to test time
                        
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", highTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                       
                    } 
                    
                }

            }
            startFHoldIndex++;
        } while (startFHoldIndex != endFirstHoldIndex);// outter for loop 
      

        //Uncomment when needing to write to the report - original
       //writeToReport(lowTcs, highTcs, failedLowTcNames, failedHighTcNames, /*dataFilePath,*/ identity);
       writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, identity, lowTimeStamp, highTimeStamp);
        
        
        findSecondHold(curespec, rowIndex, columnIndex, dataFilePath, endFirstHoldIndex);
    };
   
 public void complianceSecondHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startSHoldIndex,
            int endSecondHoldIndex) {
     ////////////////////////////////////////////////////////////////////////
     // ADDED FROM FIRST REVIEW
        List<Double> flowTcs = new ArrayList<>(); // To store tcs that did not make temp
        List<Double> fhighTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        
        List<Time> lowTimeStamp = new ArrayList<>();//to store timestamp
        List<Time> highTimeStamp = new ArrayList<>();//to store timestamp 
        
        int identity = 1;
        int dashNumber = 0;
        int startPoint = startSHoldIndex + 1;
        int locating = 0; 
        cMaxTemp = curespec.getcdMaxTemp();
        int j = 0;
     

        Sheet sheet = openExcel(dataFilePath);
        
         do {
            Row row = sheet.getRow(startSHoldIndex);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);

                    if (dashNumber == 21) {
                        dashNumber = 1;
                    } else {
                        dashNumber++; 
                    }
                    if (cell.getNumericCellValue() <= cMinTemp - lessTemp/*114.9*/) {
                        System.out.println("Low temps");
                        
                        flowTcs.add(cell.getNumericCellValue());
                       failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        
                        //was Cell cellB
                        Cell cellB = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                        Time lowTime = new Time(tcTime.getTime());//------------------------------added to test time
                        lowTimeStamp.add(lowTime);//------------------------------added to test time
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", lowTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
  
                    } else if (cell.getNumericCellValue() >= cMaxTemp + lessTemp/*145.1*/) {
                        System.out.println("High temps");
                        
                        fhighTcs.add(cell.getNumericCellValue());
                       failedHighTcNames.add(tcNames.get(dashNumber - 1)); 
                        
                       Cell cellC = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellC.getDateCellValue();//------------------------------added to test time
                        Time highTime = new Time(tcTime.getTime());//------------------------------added to test time
                        highTimeStamp.add(highTime);//------------------------------added to test time
                        
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", highTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                       
                    } 
                    
                }
            }
            startSHoldIndex++;
        } while (startSHoldIndex != endSecondHoldIndex);// outter for loop
     writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, identity, lowTimeStamp, highTimeStamp);
     findSRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
 };

     /////////////////////////////////////////////////////////////////////////////////////////////
     
     
     
     
        //List<Double> lowTcs = new ArrayList<>(); // To store tcs that did not make temp
        //List<Double> highTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        //List<String> failedLowTcNames = new ArrayList<>();
        //List<String> failedHighTcNames = new ArrayList<>();
        //int identity = 1;
        //int dashNumber = 0;
        //int startPoint = startSHoldIndex + 1;
        //cMaxTemp = curespec.getcdMaxTemp();
 

        //Sheet sheet = openExcel(dataFilePath);
       
       //do {
         //   Row row = sheet.getRow(startSHoldIndex);
           // if (row != null) {
             //   for (int item : columnIndex) {
               //     Cell cell = row.getCell(item);

                 //   if (dashNumber == 21) {
                   //     dashNumber = 1;
                   // } else {
                     //   dashNumber++;
                    //}
                    //if (cell.getNumericCellValue() <= cMinTemp - lessTemp/*239.9*/) {
       //                 lowTcs.add(cell.getNumericCellValue());
         //               failedLowTcNames.add(tcNames.get(dashNumber - 1));
           //             System.out.printf("Row: %d-%d: Low TC found: TC: %s: Temp: %.1f\n", startSHoldIndex + 1, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
             //       } else if (cell.getNumericCellValue() >= cMaxTemp + lessTemp/*265.1*/) {
               //         highTcs.add(cell.getNumericCellValue());
                 //       failedHighTcNames.add(tcNames.get(dashNumber - 1));
                   //     System.out.printf("Row: %d-%d: High TC found: TC: %s: Temp: %.1f\n", startSHoldIndex + 1, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                   // } 
               // }

            //}
            //startSHoldIndex++;
        //} while (startSHoldIndex != endSecondHoldIndex);// outter for loop 
       
        //findSRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);

        //Uncomment when needing to write to the report , I like this order more than above
        //writeToReport(lowTcs, highTcs, failedLowTcNames, failedHighTcNames, identity);
        
      //  findSRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
       
        
 //};
    
 public void complianceThirdHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startTHoldIndex,
            int endSecondHoldIndex) {

     
     List<Double> flowTcs = new ArrayList<>(); // To store tcs that did not make temp
        List<Double> fhighTcs = new ArrayList<>(); // To store tcs that exceeded max temp
        
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        
        List<Time> lowTimeStamp = new ArrayList<>();//to store timestamp
        List<Time> highTimeStamp = new ArrayList<>();//to store timestamp 
        
        int identity = 2;
        int dashNumber = 0;
        int startPoint = startTHoldIndex + 1;
        int locating = 0; 
        eMaxTemp = curespec.getefMaxTemp();
        int j = 0;
     

        Sheet sheet = openExcel(dataFilePath);
        
         do {
            Row row = sheet.getRow(startTHoldIndex);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);

                    if (dashNumber == 21) {
                        dashNumber = 1;
                    } else {
                        dashNumber++; 
                    }
                    if (cell.getNumericCellValue() <= eMinTemp - lessTemp/*114.9*/) {
                        System.out.println("Low temps");
                        
                       flowTcs.add(cell.getNumericCellValue());
                       failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        
                        //was Cell cellB
                        Cell cellB = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                        Time lowTime = new Time(tcTime.getTime());//------------------------------added to test time
                        lowTimeStamp.add(lowTime);//------------------------------added to test time
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", lowTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
  
                    } else if (cell.getNumericCellValue() >= eMaxTemp + lessTemp/*145.1*/) {
                        System.out.println("High temps");
                        
                        fhighTcs.add(cell.getNumericCellValue());
                       failedHighTcNames.add(tcNames.get(dashNumber - 1)); 
                        
                       Cell cellC = row.getCell(locating);//------------------------------added to test time
                        Date tcTime = cellC.getDateCellValue();//------------------------------added to test time
                        Time highTime = new Time(tcTime.getTime());//------------------------------added to test time
                        highTimeStamp.add(highTime);//------------------------------added to test time
                        
                        System.out.printf("\nTime: %tT, TC: %s, OOT Temp: %.1f\n", highTime, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                       
                    } 
                    
                }
            }
            startTHoldIndex++;
        } while (startTHoldIndex != endThirdHoldIndex);// outter for loop
     writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, identity, lowTimeStamp, highTimeStamp);
     //findSRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
     ////////////////////////////////////////////////////////////////////////////
     //Added to test, commented is the original
     
       // List<Double> lowTcs = new ArrayList<>(); // To store tcs that did not make temp
       // List<Double> highTcs = new ArrayList<>(); // To store tcs that exceeded max temp
       // List<String> failedLowTcNames = new ArrayList<>();
      //  List<String> failedHighTcNames = new ArrayList<>();
      //  int identity = 2;
      //  int dashNumber = 0;
      //  int startPoint = startTHoldIndex + 1;
     //   eMaxTemp = curespec.getefMaxTemp();
       

     //   Sheet sheet = openExcel(dataFilePath);
        
      //  do {
          //  Row row = sheet.getRow(startTHoldIndex);
          //  if (row != null) {
         //       for (int item : columnIndex) {
            //        Cell cell = row.getCell(item);

            //        if (dashNumber == 21) {
            //            dashNumber = 1;
            //        } else {
             ///           dashNumber++;
              //     }
             //       if (cell.getNumericCellValue() <= eMinTemp - lessTemp/*289.9*/) {
              //          lowTcs.add(cell.getNumericCellValue());
               //         failedLowTcNames.add(tcNames.get(dashNumber - 1));
              //          System.out.printf("\nRow: %d-%d: Low TC found: TC: %s: Temp: %.1f", startTHoldIndex + 1, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                //    } else if (cell.getNumericCellValue() >= eMaxTemp + lessTemp/*320.1*/) {
                //        highTcs.add(cell.getNumericCellValue());
                //        failedHighTcNames.add(tcNames.get(dashNumber - 1));
               //         System.out.printf("\nRow: %d-%d: High TC found: TC: %s: Temp: %.1f", startTHoldIndex + 1, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                //    }
               // }

            //}
            //startTHoldIndex++;
        //} while (startTHoldIndex != endThirdHoldIndex);// outter for loop 
        

        //Uncomment when needing to write to the report 
        //writeToReport(lowTcs, highTcs, failedLowTcNames, failedHighTcNames, identity);
        
       
        findTRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    };
   // writeToReport(lowTcs, highTcs, failedLowTcNames, failedHighTcNames, identity, lowTimeStamp, highTimeStamp);
   
 public void writeToReport(List<Double> flowTcs, List<Double> fhighTcs, List<String> failedLowTcNames, List<String> failedHighTcNames, int identity, List<Time> lowTimeStamp, List<Time> highTimeStamp) {

        int increaseRow;
        int reset = 1;
        int increaseColumn;
        

        try ( FileInputStream fis = new FileInputStream(reportFilePath);  Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("TempOOTs");
            
            switch(identity){
            case 0:
                
                increaseRow = 1;
                increaseColumn = 1;
               
                 Row row;
                 Cell cellA;
            int j = 0;
            for (int i = 0; i < failedLowTcNames.size(); i++) {
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String ltStamp = lowTimeStamp.get(j).toString();//well by god this worked 
                cellA.setCellValue(ltStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedLowTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(flowTcs.get(j));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow++; //= increaseRow + 1;
                j++;
            }
            //break; //purposely waterfall to next case for high temps
            case 10: 
                
                increaseRow = 1;
                increaseColumn = 8;
               j = 0;
             
            for (int i = 0; i < failedHighTcNames.size(); i++) {
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String htStamp = highTimeStamp.get(j).toString();//well by god this worked 
                cellA.setCellValue(htStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedHighTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(fhighTcs.get(j));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow++; //= increaseRow + 1;
                j++;
            }
            break;
                
            case 1:
                
                increaseRow = 1;
                increaseColumn = 15;
                j = 0;
                for (int i = 0; i < failedLowTcNames.size(); i++) {
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String ltStamp = lowTimeStamp.get(j).toString();
                cellA.setCellValue(ltStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedLowTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(flowTcs.get(j));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
                j++;
            }
            //break; //purposely waterfall to next case for high temps
            case 20: 
                
                increaseRow = 1;
                increaseColumn = 22;
               j = 0;
             
            for (int i = 0; i < failedHighTcNames.size(); i++) {
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String htStamp = highTimeStamp.get(j).toString();
                cellA.setCellValue(htStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedHighTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(fhighTcs.get(j));
                
                /*row = sheet.getRow(increaseRow);
                cellA = row.getCell(increaseColumn);
                cellA.setCellValue(failedHighTcNames.get(i));
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(fhighTcs.get(i));*/
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
                j++;
            }
            break;
            
          case 2:
                
                increaseRow = 1;
                increaseColumn = 29;
                j = 0;
                 for (int i = 0; i < failedLowTcNames.size(); i++) {
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String ltStamp = lowTimeStamp.get(j).toString();
                cellA.setCellValue(ltStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedLowTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(flowTcs.get(j));     
                     
                     
                /*row = sheet.getRow(increaseRow);
                cellA = row.getCell(increaseColumn);
                cellA.setCellValue(failedLowTcNames.get(i));
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(flowTcs.get(i));*/
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
                j++;
            }
            //break; //purposely waterfall to next case for high temps
            case 30: 
                
                increaseRow = 1;
                increaseColumn = 36;
               j = 0;
             
            for (int i = 0; i < failedHighTcNames.size(); i++) {
                
               row = sheet.getRow(increaseRow);
                cellA = row.getCell(increaseColumn);
                String htStamp = highTimeStamp.get(j).toString();
                cellA.setCellValue(htStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedHighTcNames.get(j));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(fhighTcs.get(j)); 
                
                
                
                
                /*row = sheet.getRow(increaseRow);
                cellA = row.getCell(increaseColumn);
                cellA.setCellValue(failedHighTcNames.get(i));
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(fhighTcs.get(i));*/
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
                j++;
            }
            break;     
            
            
  
        }//end switch 
 
           
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        //To test header fill
        try ( FileInputStream fis1 = new FileInputStream(reportFilePath);  Workbook workbook = new XSSFWorkbook(fis1)) {
            Sheet sheet = workbook.getSheet("Header");   
                
               int dTimeColIndex = 2;
               int dTimeRowIndex = 4;
               int fNameRowIndex = 5;
               int eqRowIndex = 6;
               int recipieRowIndex = 7;
               int jobRowIndex = 8; 
               String currDate = dateTime.toString();
               
               Row rowA;
               Cell cellA;
               
                rowA = sheet.getRow(dTimeRowIndex);
                cellA = rowA.getCell(dTimeColIndex);
                cellA.setCellValue(currDate);
                
                rowA = sheet.getRow(dTimeRowIndex + 1);
                cellA = rowA.getCell(dTimeColIndex);
                cellA.setCellValue(fileName);
                
                rowA = sheet.getRow(dTimeRowIndex + 2);
                cellA = rowA.getCell(dTimeColIndex);
                cellA.setCellValue(ovenNum);
                
                rowA = sheet.getRow(dTimeRowIndex + 3);
                cellA = rowA.getCell(dTimeColIndex);
                cellA.setCellValue(runRecipe);
                
                rowA = sheet.getRow(dTimeRowIndex + 4);
                cellA = rowA.getCell(dTimeColIndex);
                cellA.setCellValue(cureJob);
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                }
        
          }catch (IOException e) {
            e.printStackTrace();
        }

        
       // Desktop.getDesktop().open(new File(reportFilePath));
        
    }
                                        //ramp rates                           //TcNames 
 public void writeRampToReport( List<Double> rampRateList, int identity, List<String> failedTcNames, List<Time> timeStamp){
 
 try ( FileInputStream fis = new FileInputStream(reportFilePath);  Workbook workbook = new XSSFWorkbook(fis)) {
            
        int increaseRow;
        int reset = 1;
        int increaseColumn;
        
     
     
     
     Sheet sheet = workbook.getSheet("RampRateOOTs");
            
            switch(identity){
            case 3:
                increaseRow = 1;
                increaseColumn = 1;
               
                 Row row;
                 Cell cellA;

            for (int i = 0; i < failedTcNames.size(); i++) {
                
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String tStamp = timeStamp.get(i).toString();//well by god this worked 
                cellA.setCellValue(tStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedTcNames.get(i));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(rampRateList.get(i));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
            }
            break;
            
            case 4:
                increaseRow = 1;
                increaseColumn = 8;
               
                 //Row row;
                 //Cell cellA;

            for (int i = 0; i < failedTcNames.size(); i++) {
                
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String tStamp = timeStamp.get(i).toString();//well by god this worked 
                cellA.setCellValue(tStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedTcNames.get(i));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(rampRateList.get(i));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
            }
            break;
            
            case 5:
                increaseRow = 1;
                increaseColumn = 15;
               
                 //Row row;
                 //Cell cellA;

            for (int i = 0; i < failedTcNames.size(); i++) {
                
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String tStamp = timeStamp.get(i).toString();//well by god this worked 
                cellA.setCellValue(tStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedTcNames.get(i));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(rampRateList.get(i));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
            }
            break;
            
            case 6:
                increaseRow = 1;
                increaseColumn = 22;
               
                 //Row row;
                 //Cell cellA;

            for (int i = 0; i < failedTcNames.size(); i++) {
                
                row = sheet.getRow(increaseRow);
                
                cellA = row.getCell(increaseColumn);
                String tStamp = timeStamp.get(i).toString();//well by god this worked 
                cellA.setCellValue(tStamp);
                
                cellA = row.getCell(increaseColumn + 1);
                cellA.setCellValue(failedTcNames.get(i));
                
                cellA = row.getCell(increaseColumn + 2);
                cellA.setCellValue(rampRateList.get(i));
                
                try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                workbook.write(fos);
                }catch (IOException e){
                    e.printStackTrace();
                };
                increaseRow = increaseRow + 1;
            }
            break;

            }
            
 }catch (IOException e) {
            e.printStackTrace();
        }
 };

 
 
 public void findRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int holdIndex,
            int endHoldIndex, Sheet sheet) {

        //int increase = 0;
        double currentNum = 0;
        double previousNum = 0;
        int identity = 3;
       // int item = 4;
       int j= 0;//------------------------------added to test time
        int startingPoint = rowIndex + 2;
        int nextCol = startingPoint;
        int dashNumber = 0;
        rampRateMax = curespec.getRampRate();
        List<String> failedRampTcs = new ArrayList<>();
        List<Double> columnDeltas = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();//------------------------------added to test time
        int locating = 0;//------------------------------added to test time
        System.out.println("\n**********************\n");
        System.out.println("\nFirst Ramp OOTs\n");
   

            for (int item1 : columnIndex) { //swap place with do loop maybe 
                do {
                
                Row row = sheet.getRow(holdIndex);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    
         
                    if(holdIndex == 25){
                        previousNum = currentNum;
                    }
                    if(dashNumber == 21){
                        dashNumber = 1;
                    }
                    
                    double rampRate = Math.abs(currentNum - previousNum);
 
                    if (rampRate >= rampRateMax){
                    columnDeltas.add(rampRate);
                    failedRampTcs.add(tcNames.get(dashNumber));
                    //System.out.printf("\nCurrent Column: %d, Current num: %f, Previous Num %f, OOT Ramp Rate: %.1f, TC: %s\n", item1, currentNum, previousNum, rampRate, tcNames.get(dashNumber));// lets view the calculation
                    
                    Cell cellB = row.getCell(locating);//------------------------------added to test time
                    Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                    Time time = new Time(tcTime.getTime());//------------------------------added to test time
                    timeStamp.add(time);//------------------------------added to test time
                    System.out.printf("\nTime: %tT, TC: %s, OOT Ramp Rate: %.1f, Current num: %f, Previous Num %f\n", time, tcNames.get(dashNumber), rampRate, currentNum, previousNum);

                    }
   
                }
                previousNum = currentNum;
                holdIndex++; //testing
            }while (holdIndex != endHoldIndex);
            dashNumber++;
            holdIndex = nextCol;
        } 
        System.out.println("\n**********************\n");
        writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);//------------------------------commented to test time
    };
 
    ////Original MEthod 
   /*public void findRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int holdIndex,
            int endHoldIndex, Sheet sheet) {

        //int increase = 0;
        double currentNum = 0;
        double previousNum = 0;
        int identity = 3;
       // int item = 4;
        int startingPoint = rowIndex + 2;
        int nextCol = startingPoint;
        int dashNumber = 0;
        rampRateMax = curespec.getRampRate();
        List<String> failedRampTcs = new ArrayList<>();
        List<Double> columnDeltas = new ArrayList<>();
        System.out.println("\n**********************\n");
        System.out.println("\nFirst Ramp OOTs\n");
   

            for (int item1 : columnIndex) { //swap place with do loop maybe 
                do {
                
                Row row = sheet.getRow(holdIndex);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
         
                    if(holdIndex == 25){
                        previousNum = currentNum;
                    }
                    if(dashNumber == 21){
                        dashNumber = 1;
                    }
                    
                    double rampRate = Math.abs(currentNum - previousNum);
 
                    if (rampRate >= rampRateMax){
                    columnDeltas.add(rampRate);
                    failedRampTcs.add(tcNames.get(dashNumber));
                    System.out.printf("\nCurrent Column: %d, Current num: %f, Previous Num %f, OOT Ramp Rate: %.1f, TC: %s\n", item1, currentNum, previousNum, rampRate, tcNames.get(dashNumber));// lets view the calculation

                    }
   
                }
                previousNum = currentNum;
                holdIndex++; //testing
            }while (holdIndex != endHoldIndex);
            dashNumber++;
            holdIndex = nextCol;
        } 
        System.out.println("\n**********************\n");
        writeRampToReport(columnDeltas, identity, failedRampTcs);//Pray to jesus lol
    };*/

    
    //Find second ramp rate
 public void findSRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        int increase = 0;
        double currentNum = 0;
        double previousNum = 0;
        int identity = 4;
        int item = 4;
        int j = 0;
        int locating = 0;
        int dashNumber = 0;

        int startingPoint = endFirstHoldIndex + 1;
        int endingPoint = startSecondHoldIndex;
        int nextCol = startingPoint;
        List<String> failedRampTcs = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        List<Double> columnDeltas = new ArrayList<>();
        System.out.println("\n**********************\n");
        System.out.println("\nSecond Ramp OOTs\n");

            for (int item1 : columnIndex) {
                
                do {
                    
                Row row = sheet.getRow(startingPoint);
                if (row != null) {
                    
                    
                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    
                     if(startingPoint == nextCol){
                        previousNum = currentNum;
                        
                    }
                     if(dashNumber == 21){
                        dashNumber = 1;
                    }
                     
                    double rampRate = Math.abs(currentNum - previousNum);

                    if (rampRate >= rampRateMax){
                    columnDeltas.add(rampRate);
                    failedRampTcs.add(tcNames.get(dashNumber));
                    
                    Cell cellB = row.getCell(locating);//------------------------------added to test time
                    Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                    Time time = new Time(tcTime.getTime());//------------------------------added to test time
                    timeStamp.add(time);//------------------------------added to test time
                    System.out.printf("\nTime: %tT, TC: %s, OOT Ramp Rate: %.1f, Current num: %f, Previous Num %f\n", time, tcNames.get(dashNumber), rampRate, currentNum, previousNum);

                    //System.out.printf("\nCurrent Column: %d, Current num: %f, Previous Num %f, Ramp Rate: %.1f\n", item1, currentNum, previousNum, rampRate);// lets view the calculation
                   
                    }

                } 
                previousNum = currentNum;
                startingPoint++; 

          
        } while (startingPoint != endingPoint);//if this works you need to make it automatically find the number of TCs and set the condition to that
                startingPoint = nextCol; 
                dashNumber++;
     //
    }
            System.out.println("\n**********************\n");
            writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);//------------------------------commented to test time
 
 };
    
    
    //Find the thrid ramp rate 
 public void findTRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        int increase = 0;
        double currentNum = 0;
        double previousNum = 0;
        int identity = 5;
        int item = 4;
        int dashNumber = 0;
        int locating = 0;

        int startingPoint = endSecondHoldIndex + 1;
        int endingPoint = startThirdHoldIndex;
        int nextCol = startingPoint;

        //int startingPoint = rowIndex + 2; 
        List<Double> columnDeltas = new ArrayList<>();
        List<String> failedRampTcs = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        
        
        
        //do {
        System.out.println("\n**********************\n");
        System.out.println("\nThird Ramp OOTs\n");
            for (int item1 : columnIndex) {
                
                do {
                Row row = sheet.getRow(startingPoint);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    
                    
                    if(startingPoint == nextCol){
                        previousNum = currentNum;
                    }
                    if(dashNumber == 21){
                        dashNumber = 1;
                    }

                    
                    double rampRate = Math.abs(currentNum - previousNum);
                    if (rampRate >= rampRateMax){
                    columnDeltas.add(rampRate);
                    failedRampTcs.add(tcNames.get(dashNumber));
                    Cell cellB = row.getCell(locating);//------------------------------added to test time
                    Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                    Time time = new Time(tcTime.getTime());//------------------------------added to test time
                    timeStamp.add(time);//------------------------------added to test time
                    System.out.printf("\nTime: %tT, TC: %s, OOT Ramp Rate: %.1f, Current num: %f, Previous Num %f\n", time, tcNames.get(dashNumber), rampRate, currentNum, previousNum);
                    //System.out.printf("\nCurrent Column: %d, Current num: %f, Previous Num %f, Ramp Rate: %.1f\n", item1, currentNum, previousNum, rampRate);// lets view the calculation
  
                    }

                }
                previousNum = currentNum;
                startingPoint++;
            //}
            //startingPoint++;
        } while (startingPoint != endingPoint);//if this works you need to make it automatically find the number of TCs and set the condition to that
        startingPoint = nextCol;
        dashNumber++;
        
    }
            System.out.println("\n**********************\n");
            writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);//------------------------------commented to test time
       findCoolRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
 };
    
   //Works 
 public void findCoolRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        int increase = 0;
        double currentNum = 0;
        double previousNum = 0;
        int identity = 6;
        int item = 4;
        int dashNumber = 0;
        int locating = 0;
        int startingPoint = endThirdHoldIndex;
        int endingPoint = 0;
        int nextCol = startingPoint;
        List<String> failedRampTcs = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        List<Double> columnDeltas = new ArrayList<>();
        
        System.out.println("\n**********************\n");
        System.out.println("\nCooling Ramp OOTs\n");
        //do {

            for (int item1 : columnIndex) {
                do {
                Row row = sheet.getRow(startingPoint);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    
                    
                    if(startingPoint == nextCol){
                        previousNum = currentNum;
                    }
                    if(dashNumber == 21){
                        dashNumber = 1;
                    }

                    
                    double rampRate = Math.abs(previousNum - currentNum);
                    if (rampRate >= rampRateMax){
                    columnDeltas.add(rampRate);
                    failedRampTcs.add(tcNames.get(dashNumber));
                    Cell cellB = row.getCell(locating);//------------------------------added to test time
                    Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                    Time time = new Time(tcTime.getTime());//------------------------------added to test time
                    timeStamp.add(time);//------------------------------added to test time
                    System.out.printf("\nTime: %tT, TC: %s, OOT Ramp Rate: %.1f, Current num: %f, Previous Num %f\n", time, tcNames.get(dashNumber), rampRate, currentNum, previousNum);
                    //System.out.printf("\nCurrent Column: %d, Current num: %f, Previous Num %f, Ramp Rate: %.1f\n", item1, currentNum, previousNum, rampRate);// lets view the calculation
                    //System.out.println("Start hold:" + startingPoint);
                    //System.out.println("End Hold: " + endingPoint);
                    }
            
                }
                //This is kind of ugly determine a new way later on
                if (row == null) {
                    endingPoint = startingPoint;
                    //System.out.printf("StartingPoint: %d, EndingPoint %d", startingPoint, endingPoint);

                    break;
                }
                previousNum = currentNum;
                startingPoint++;
                endingPoint++;

            //}
            //startingPoint++;
            //endingPoint++;
        } while (startingPoint != endingPoint);//if this works you need to make it automatically find the number of TCs and set the condition to that
                startingPoint = nextCol;
                dashNumber++;
        dataEnd = endingPoint;
        //findVacuum(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    }
            System.out.println("\n**********************\n");
            writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);//------------------------------commented to test time
            findVacuum(curespec, rowIndex, columnIndex, dataFilePath, sheet);
 };

    
    ///Does work
 public void findVacuum(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
        System.out.println("\n**********************\n");
        System.out.println("\nVacuum Pressure OOTs\n");
        int dashNum = 24;
        int dashNumber = 0;
        int identity = 7;
        vacPressure = curespec.getVacuum();
        minVacPressure = vacPressure - 6;
        List<String> failedVacInHg = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        int locating = 0; 
        

        do {
            Row row = sheet.getRow(dataStart);
            if (row != null) {
                for (int item : vacuumIndex) {
                    Cell cell = row.getCell(item);

                    if (dashNum == 33) {
                        dashNum = 25;
                    } else {
                        dashNum++;
                    }
                    if (dashNumber == 9){
                        dashNumber = 0;
                      }
                    else {
                        dashNumber++;
                    }
                  
                     
                    double vacPres = Math.abs(cell.getNumericCellValue());
                    
                    
                    //Math.abs()
                    if (Math.abs(cell.getNumericCellValue()) <= minVacPressure/*23.9*/) {
                    
                        //vacInHg.add(vacPres);
                        //failedVacInHg.add(vacNames.get(dashNumber));
                        //System.out.println(failedVacInHg.get(dashNumber));
                        
                    Cell cellB = row.getCell(locating);//------------------------------added to test time
                    Date tcTime = cellB.getDateCellValue();//------------------------------added to test time
                    Time time = new Time(tcTime.getTime());//------------------------------added to test time
                    timeStamp.add(time);//------------------------------added to test time
                    System.out.printf("\nTime: %tT, VAC-Pressure: %.1f", time, vacPres);
                   
                    //dashNumber++;
                    
                  // System.out.printf("\nTime: %tT, TC: %s, OOT Ramp Rate: %.1f, Current num: %f, Previous Num %f\n", time, tcNames.get(dashNumber), rampRate, currentNum, previousNum);   
                    //lowTcs.add(cell.getNumericCellValue());
                        //failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        //System.out.printf("\nTime: %tT, Row: %d-%d: Low VAC-TC found: VAC-Pressure: %.1f", time, dataStart + 1, dashNum, cell.getNumericCellValue());
                    } //else if (Math.abs(cell.getNumericCellValue()) >= vacPressure/*29.9*/) {
                        //highTcs.add(cell.getNumericCellValue());
                        //failedHighTcNames.add(tcNames.get(dashNumber - 1));
                        //System.out.printf("\nRow: %d-%d: High TC found: TC: %s: Temp: %.1f", startTHoldIndex + 1, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                        //System.out.printf("\nRow: %d-%d: High VAC-TC found: VAC-Pressure: %.1f", dataStart + 1, dashNum, cell.getNumericCellValue());
                   // } else {
                      //  System.out.printf("\nRow: %d-%d Tc's passed VAC requirements...", dataStart + 1, dashNum);
                    //}
                }

            }
            dataStart++;
           
           // dashNumber++;
        } while (dataStart != dataEnd);

        //dashNumber++;
        dataStart = 0;
        System.out.println("\n**********************\n");
        
        //Testing Vac TC Names
       //for(int i = 0; i < vacNames.size(); i++){
         //  System.out.println(vacNames.get(i));
        //}

        find9002(curespec, rowIndex, columnIndex, dataFilePath, sheet);

    }

    ;
 
 
 //Works as well 
 public void find9002(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        System.out.println("\n**********************\n");
        System.out.println("\n9002 Temp OOTs\n");
        //System.out.println(dataStart2);
        int identity = 8;
        int dashNumber = 0;
        temp9002 = curespec.getTemp9002();

        //sheet = openExcel(dataFilePath);
        //for(int ind : columnIndex) {
        do {
            Row row = sheet.getRow(dataStart2);
            if (row != null) {
                for (int item : columnIndex) {
                    Cell cell = row.getCell(item);

                    if (dashNumber == 21) {
                        dashNumber = 1;
                    } else {
                        dashNumber++;
                    }
                    if (cell.getNumericCellValue() >= temp9002/*360.1*/) {
                        //lowTcs.add(cell.getNumericCellValue());
                        //failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        // System.out.printf("\nRow: %d-%d: Low VAC-TC found: VAC-Pressure: %.1f", dataStart + 1, dashNum, cell.getNumericCellValue());
                        System.out.printf("\nRow: %d-%d: 9002 TC found: TC Temp: %.1f", dataStart2 + 1, dashNumber, cell.getNumericCellValue());
                    } //else if (cell.getNumericCellValue() >= 145.1) {
                    //highTcs.add(cell.getNumericCellValue());
                    //failedHighTcNames.add(tcNames.get(dashNumber - 1));
                    //System.out.printf("\nRow: %d-%d: High TC found: TC: %s: Temp: %.1f", startFHoldIndex, dashNumber, tcNames.get(dashNumber - 1), cell.getNumericCellValue());
                   // else {
                     //   System.out.printf("\nRow: %d-%d Tc's passed 9002 requirements...", dataStart2 + 1, dashNumber);
                    //}
                }

            }
            dataStart2++;
        } while (dataStart2 != dataEnd);
        System.out.println("\n**********************\n");
        //Desktop.getDesktop().open(new File(reportFilePath));
        //getTimeStamp(curespec,rowIndex, columnIndex, dataFilePath, sheet);
    }
;

 
 public void getTimeStamp(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
 
     int locating = 0;
     boolean condition = true;
     int newRowIndex = rowIndex + 2;
     
     List<Time> timeStamp = new ArrayList<>();
     //var .add(array.toString()
 
        sheet = openExcel(dataFilePath); 
       
        do {
            Row row = sheet.getRow(newRowIndex);
            if (row != null) {
            
                //for (int item : columnIndex) {
                   Cell cell = row.getCell(locating);
                   Date tcTime = cell.getDateCellValue();
                   Time time = new Time(tcTime.getTime());
                   
                   //String tcTimeStr = .toString(tcTime);
                   timeStamp.add(time);
                   System.out.println(time);
            }
            else{
                condition = false;
            }
            newRowIndex++;

            }while(condition = true);
 }
 
 
} //End Subclass ExcelHandler


