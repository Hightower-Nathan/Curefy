package CurefyPkg;

/**
 * @Author Name: Nathan Hightower
 * @Project Name: Curefy
 * @Date: Feb 8, 2025
 * @Subclass excelHandler Description: This currently includes all of the methods
 * for every bit of data that needs to be reviewed as well as those designed to 
 * generate the report 
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

//Begin Subclass ExcelHandler
public class ExcelHandler extends CureSpec {

    //Class Constructors 
    public ExcelHandler(){};
     
    public ExcelHandler(CureSpec curespec){};
    
    
    //Variables for class to work with excel data
    int startFirstHoldIndex, startSecondHoldIndex, startThirdHoldIndex;
    int endFirstHoldIndex, endSecondHoldIndex, endThirdHoldIndex;
    List<String> tcNames = new ArrayList<>();
    List<Integer> vacuumIndex = new ArrayList<>();
    List<String> vacNames = new ArrayList<>();
    List<Double> vacInHg = new ArrayList<>();
    int dataStart = 0;
    int dataEnd = 0;
    int dataStart2 = 0;
    private String fileName, ovenNum, runRecipe, cureJob;
    //Private Variables from curespec
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
    //Random Variables to hold report path and create timestamps for report
    Date dateTime = new Date();
    
    //Change to your desired filepath 
    String reportFilePath = 
       "C:/Program Files/NetBeans-12.6/WorkSpace/CurefyReportTemplate(9).xlsx";
   

    /**
     * Method findRowIndex: This works to find where in the datafile the 
     * TC data is located. Once found it stores that rowIndex for later use. 
     * It also takes the time to store the dataFile header information 
     * for use in the report file generated at the end of the run
     * @param dataFilePath
     * @return 
     */
    public int findRowIndex(String dataFilePath) {
        int rowIndex = 0;
        Sheet sheet = openExcel(dataFilePath);

        for (Row row : sheet) {
            for (Cell cell : row) {
                //Start of TC data always start where the first column
                //mentions "Time" look for and start at that index + 2
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Time")) {
                    rowIndex = row.getRowNum();
                    // to be used in the vacuum verify portion 
                    dataStart = rowIndex + 2;
                    dataStart2 = rowIndex + 2;
                }
                
                //Header portion here .contains looks for strings and stores 
                //those values at those indexes
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Filename:")){
                    Cell nextTo = row.getCell(1);
                    fileName = nextTo.toString();
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Equipment:")){
                    Cell nextTo = row.getCell(1);
                    ovenNum = nextTo.toString();
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Run Recipe:")){
                    Cell nextTo = row.getCell(1);
                    runRecipe = nextTo.toString();
                }
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("Part Number:")){
                    Cell nextTo = row.getCell(1);
                    cureJob = nextTo.toString();
                }
                break;
            }
        }
        return rowIndex;
    }; // End method findRowIndex
    
    /**
     * Method findColumnIndex: This looks by column to find the TC names for 
     * tempature TC and for Vacuum TC to be used in report using the located 
     * row index. It stores these in two arrays to give us our boundry of 
     * data within the data file or what ranges to look between for review
     * @param dataFilePath
     * @param rowIndex
     * @return 
     */
    public List<Integer> findColumnIndex(String dataFilePath, int rowIndex) {
        int columnIndexx = 0;
        int vacuumIndexx = 0;
        List<Integer> columnIndex = new ArrayList<>();
        Sheet sheet = openExcel(dataFilePath);
        Row row = sheet.getRow(rowIndex);

        if (row != null) {
            for (Cell cell : row) {
                //Store found PTCs in the columnIndex array
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("PTC")) {
                    columnIndexx = cell.getColumnIndex();
                    columnIndex.add(columnIndexx);
                    //Store found VPRB in the vaccumIndex array
                } else if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("VPRB")) {
                    vacuumIndexx = cell.getColumnIndex();
                    vacuumIndex.add(vacuumIndexx);
                    vacNames.add(cell.getStringCellValue()); //To store the names of the vacuum TCS for pressure check.  
                }
            }
        }
        return columnIndex;
    }; //End method findColumnIndex
   
    /**
     * Method readTcNames: This stores the names of the TCs where the cells 
     * contain the value PTC
     * @param dataFilePath
     * @param rowIndex
     * @return 
     */
    public List<String> readTcNames(String dataFilePath, int rowIndex) {

        Sheet sheet = openExcel(dataFilePath);
        Row row = sheet.getRow(rowIndex);

        //This is an often construct in this project. It means if the row is not 
        //empty and there is a cell in the row and if that cell contains the 
        //string PTC store that value as the TC name inside of the tcNames array
        if (row != null) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING
                        && cell.getStringCellValue().contains("PTC")) {
                    tcNames.add(cell.getStringCellValue());
                }
            }
        }
        return tcNames;
    }; //End method readTcNames
    
    /**
     * Method: Reads the data on column 1 of the specified cure spec
     * This is the designated column that has the cure spec required values
     * which are read and stored for the reviewing process 
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
    }//End method  readColumn

    /**
     * Method findFirstHold: This looks for the first hold by identifying where 
     * the lagging TC reaches the minimal temp for the hold. Once found it sends 
     * it the method that locates the end of that hold. By doing this it will
     * block out the first hold so the review can begin. 
     * @param curespec
     * @param rowIndex
     * @param columnIndex
     * @param dataFilePath 
     */
    public void findFirstHold(CureSpec curespec, int rowIndex, 
            List<Integer> columnIndex, String dataFilePath) {
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
                //As long as there is an element in the columnIndex array - loop
                for (int item : columnIndex) {
                    //Get the item at that cell index
                    Cell cell = row.getCell(item);
                    //Set current temp to condition and compare to min temp requirement 
                    if (cell.getNumericCellValue() <= aMinTemp) {
                        condition = cell.getNumericCellValue();
                        //If the condition is greater than the minTemp for all
                        //Columns grab the index and assign it to startFirstHoldIndex
                        if (condition >= aMinTemp) {
                            startFHoldIndex = row.getRowNum();
                            startFirstHoldIndex = startFHoldIndex;
                        } else {
                            break;
                        }
                    }
                }
            }
            //Increment to jump to the next row in the datafile
            startRowInd++;

        } while (condition <= stopCondition);// outter for loop 
        startRowInd = rowIndex + 2;
        //Method call to find the ramp rates of the data between data start and first hold start index 
        findRampRate(curespec, rowIndex, columnIndex, dataFilePath, startRowInd,
                startFirstHoldIndex, sheet);

        //Method call to find the end of the first hold 
        findendFirstHold(curespec, dataFilePath, startFHoldIndex, columnIndex, 
                rowIndex);

    }; // End method findFirstHold 
    
    /**
     * Method findEndFirstHold as above this finds the place where the first hold ends. 
     * Once found it will supply the start/end indexes to the compliance method 
     * for review 
     * @param curespec
     * @param datafilePath
     * @param startFirstHoldIndex
     * @param columnIndex
     * @param rowIndex 
     */
   public void findendFirstHold(CureSpec curespec, String datafilePath, 
           int startFirstHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endFHoldIndex = 0;
        //Method call to retrieve the step b cure hold requirement from the curespec class 
        double bHoldTimeMinutes = curespec.getbHoldTime();
        //Counter for elapsed minutes
        double elapsedMinutes = 0.0;
        int originalValue = startFirstHoldIndex + 1;
        //Open dataFile 
        Sheet sheet = openExcel(datafilePath);

        do {
            //Begin at the first hold index previously identified 
            Row row = sheet.getRow(startFirstHoldIndex);
            //Search that row in all columnIndex 
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endFHoldIndex = row.getRowNum();
                }

                //IF the elapsed minutes met the required hold number of minutes break out 
                //of the loop 
                if (elapsedMinutes == bHoldTimeMinutes) {
                    break;
                }
            }
            //Increment number of minutes and index position 
            elapsedMinutes++;
            startFirstHoldIndex++;
        } while (elapsedMinutes != bHoldTimeMinutes + 1);

        
        endFirstHoldIndex = endFHoldIndex;
        //Method call to preform the first hold review 
        complianceFirstHold(curespec, rowIndex, columnIndex, datafilePath, 
                originalValue, endFirstHoldIndex);

    };//End method 
     
   /**
    * Method findSecondHold this does exactly what the find firstFirstHold
    * method does. The only difference is that it uses the ending index of the 
    * previous hold to determine the start position. 
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param endFirstHoldIndex 
    */
   public void findSecondHold(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, 
           int endFirstHoldIndex) {
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

        //Method call to find the end of the second hold
        findendSecondHold(curespec, dataFilePath, startSHoldIndex, columnIndex, 
                rowIndex);
    }; // End Method findSecondHold 
   
   /**
    * Method findEndSecondHold finds the index where the second hold ends. 
    * Once found it supplies it to the method for reviewing the hold as well
    * as to the method for finding the third hold location
    * @param curespec
    * @param datafilePath
    * @param startSecondHoldIndex
    * @param columnIndex
    * @param rowIndex 
    */
   public void findendSecondHold(CureSpec curespec, String datafilePath, 
           int startSecondHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endSHoldIndex = 0;
        double dHoldTimeMinutes = curespec.getdHoldTime();
        double elapsedMinutes = 0.0;
        int originalValue = startSecondHoldIndex + 1;

        Sheet sheet = openExcel(datafilePath);
        do {
            //Start at the starting hold index and search until the number of 
            //elapsed minutes is equal to the required number of minutes
            //Once equal store that row index
            Row row = sheet.getRow(startSecondHoldIndex);
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endSHoldIndex = row.getRowNum();
                }

                if (elapsedMinutes == dHoldTimeMinutes) {
                    break;
                }
            }
            elapsedMinutes++;
            startSecondHoldIndex++;
        } while (elapsedMinutes != dHoldTimeMinutes + 1);
        endSecondHoldIndex = endSHoldIndex;
        
        //Method call to review the identified second hold 
        complianceSecondHold(curespec, rowIndex, columnIndex, datafilePath, 
                originalValue, endSecondHoldIndex);

        //Method call to find the starting third hold index
        findThirdHold(curespec, rowIndex, columnIndex, datafilePath, 
                endSecondHoldIndex);
    }; // End method findEndSecondHold 
   
   /**
    * Method findThirdHold: finds the starting third old index 
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param endSecondHoldIndex 
    */
   public void findThirdHold(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, 
           int endSecondHoldIndex) {
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

        //Method call to find the ending index of the third hold 
        findendThirdHold(curespec, dataFilePath, startTHoldIndex, columnIndex, rowIndex);
    };// End method findThirdHold 
   
   /**
    * Method findEndThirdHold: Finds the ending index for the third hold which 
    * will be supplied to the third review 
    * @param curespec
    * @param datafilePath
    * @param startThirdHoldIndex
    * @param columnIndex
    * @param rowIndex 
    */
   public void findendThirdHold(CureSpec curespec, String datafilePath, 
           int startThirdHoldIndex, List<Integer> columnIndex, int rowIndex) {
        int endTHoldIndex = 0;
        double fHoldTimeMinutes = curespec.getfHoldTime();
        double elapsedMinutes = 0.0;
        int originalValue = startThirdHoldIndex + 1;
        Sheet sheet = openExcel(datafilePath);
        do {
            Row row = sheet.getRow(startThirdHoldIndex);
            for (int item : columnIndex) {
                Cell cell = row.getCell(item);
                if (cell != null) {
                    endTHoldIndex = row.getRowNum();
                }

                if (elapsedMinutes == fHoldTimeMinutes) {
                    break;
                }
            }
            elapsedMinutes++;
            startThirdHoldIndex++;
        } while (elapsedMinutes != fHoldTimeMinutes + 1);

        endThirdHoldIndex = endTHoldIndex;
        //Method call to review the third hold 
        complianceThirdHold(curespec, rowIndex, columnIndex, datafilePath, originalValue,//was startFirstHoldIndex
                endFirstHoldIndex);
    }; // End Method findEndThirdHold
   
   /**
    * Method openExcel takes the datafile path and opens it for review 
    * @param dataFilePath
    * @return 
    */
   public Sheet openExcel(String dataFilePath) {

        Sheet s = null;
        try ( FileInputStream fis2 = new FileInputStream(dataFilePath);  
                Workbook workbook = new XSSFWorkbook(fis2)) {
            Sheet sheet = workbook.getSheetAt(0);
            s = sheet;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return s;
    }; // End Method openExcel
   
/**
 * Method complianceFirstHold: This method will review all cells within the 
 * starting and ending hold indexes. It will look at each cell value comparing
 * to the value required by the curespec. If it fails to meet that value it will
 * store that value in the array for failed low/high TCs. It will also store
 * the TC name as well as the timestamp
 * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
 * @param curespec
 * @param rowIndex
 * @param columnIndex
 * @param dataFilePath
 * @param startFHoldIndex
 * @param endFirstHoldIndex 
 */
 public void complianceFirstHold(CureSpec curespec, int rowIndex, 
         List<Integer> columnIndex, String dataFilePath, int startFHoldIndex,
            int endFirstHoldIndex) {
        System.out.println("First Hold");
        // To store tcs that did not make temp
        List<Double> flowTcs = new ArrayList<>(); 
        // To store tcs that exceeded max temp
        List<Double> fhighTcs = new ArrayList<>(); 
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        //to store timestamp
        List<Time> lowTimeStamp = new ArrayList<>();
        //to store timestamp 
        List<Time> highTimeStamp = new ArrayList<>();
        int identity = 0;
        int dashNumber = 0;
        int startPoint = startFHoldIndex;
        int locating = 0;
        //Get the max temp value from the cure spec
        aMaxTemp = curespec.getabMaxTemp();
        int j = 0;
        Sheet sheet = openExcel(dataFilePath);

        do {
            Row row = sheet.getRow(startFHoldIndex);
            if (row != null) {
                //While there is an element in the columnIndex array
                for (int item : columnIndex) {
                    //Get that element 
                    Cell cell = row.getCell(item);

                    //This is meant to use as an index when pulling the TC
                    // names for locating the correct name that is to be paired
                    // with the correct data. If dashNumber is 21 (21 TCs then 
                    // reset) else increment 
                    if (dashNumber == 21) {//This block is repeated make a singular function ************************************** For future updates
                        dashNumber = 1;
                    } else {
                        dashNumber++;
                    }
                    //If the TC has a value less than the min required per spec
                    // store that TC name, TC data and time stamp
                    if (cell.getNumericCellValue() <= aMinTemp - 
                            lessTemp/*114.9*/) {
                        flowTcs.add(cell.getNumericCellValue());
                        failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates 
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time lowTime = new Time(tcTime.getTime());
                        lowTimeStamp.add(lowTime);

                    //If the TC has a value greater than the max required per spec
                    // store that TC name, TC data and time stamp
                    } else if (cell.getNumericCellValue() >= aMaxTemp + 
                            lessTemp/*145.1*/) {
                        fhighTcs.add(cell.getNumericCellValue());
                        failedHighTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates 
                        Cell cellC = row.getCell(locating);
                        Date tcTime = cellC.getDateCellValue();
                        Time highTime = new Time(tcTime.getTime());
                        highTimeStamp.add(highTime);
                    }
                }
            }
            // Jump to the next row in the data file 
            startFHoldIndex++;
            // Loop while the starting index hasnt reached the ending index 
        } while (startFHoldIndex != endFirstHoldIndex);// outter for loop 

        // Method call to Write the descrepant TCs to the report 
        writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, 
                identity, lowTimeStamp, highTimeStamp);
        // Method call to find the second hold data
        findSecondHold(curespec, rowIndex, columnIndex, dataFilePath, 
                endFirstHoldIndex);
    }; // End method complianceFirstHold 
   
   /**
    * Method complianceSecondHold: This method will review all cells within the
    * starting and ending hold indexes. It will look at each cell value comparing
    * to the value required by the curespec. If it fails to meet that value it will
    * store that value in the array for failed low/high TCs. It will also store
    * the TC name as well as the timestamp
    * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param startSHoldIndex
    * @param endSecondHoldIndex 
    */
   public void complianceSecondHold(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, int startSHoldIndex,
            int endSecondHoldIndex) {
        System.out.println("Second Hold");
        // To store tcs that did not make temp
        List<Double> flowTcs = new ArrayList<>(); 
        // To store tcs that exceeded max temp
        List<Double> fhighTcs = new ArrayList<>(); 
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        //to store timestamp
        List<Time> lowTimeStamp = new ArrayList<>();
        //to store timestamp 
        List<Time> highTimeStamp = new ArrayList<>();
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

                    //This is meant to use as an index when pulling the TC
                    // names for locating the correct name that is to be paired
                    // with the correct data. If dashNumber is 21 (21 TCs then 
                    // reset) else increment 
                    if (dashNumber == 21) {//This block is repeated make a singular function ************************************** For future updates
                        dashNumber = 1;
                    } else {
                        dashNumber++;
                    }
                    //If the TC has a value less than the min required per spec
                    // store that TC name, TC data and time stamp
                    if (cell.getNumericCellValue() <= cMinTemp - 
                            lessTemp/*114.9*/) {
                        flowTcs.add(cell.getNumericCellValue());
                        failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time lowTime = new Time(tcTime.getTime());
                        lowTimeStamp.add(lowTime);

                    //If the TC has a value greater than the max required per spec
                    // store that TC name, TC data and time stamp
                    } else if (cell.getNumericCellValue() >= cMaxTemp + 
                            lessTemp/*145.1*/) {
                        fhighTcs.add(cell.getNumericCellValue());
                        failedHighTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellC = row.getCell(locating);
                        Date tcTime = cellC.getDateCellValue();
                        Time highTime = new Time(tcTime.getTime());
                        highTimeStamp.add(highTime);
                    }
                }
            }
            // Jump to the next row in the data file 
            startSHoldIndex++;
            // Loop while the starting index hasnt reached the ending index
        } while (startSHoldIndex != endSecondHoldIndex);// outter for loop
        // Method call to Write the descrepant TCs to the report
        writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, 
                identity, lowTimeStamp, highTimeStamp);
        
        // Method call to find the ramp rates between the first/second hold
        findSRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    }; // End method complianceSecondHold 
    
   /**
    * Method complianceThirdHold: This method will review all cells within the 
    * starting and ending hold indexes. It will look at each cell value comparing 
    * to the value required by the curespec. If it fails to meet that value it will
    * store that value in the array for failed low/high TCs. It will also store
    * the TC name as well as the timestamp
    * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param startTHoldIndex
    * @param endSecondHoldIndex 
    */
   public void complianceThirdHold(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, int startTHoldIndex,
            int endSecondHoldIndex) {
        System.out.println("Third hold");
        // To store tcs that did not make temp
        List<Double> flowTcs = new ArrayList<>(); 
        // To store tcs that exceeded max temp
        List<Double> fhighTcs = new ArrayList<>(); 
        List<String> failedLowTcNames = new ArrayList<>();
        List<String> failedHighTcNames = new ArrayList<>();
        //to store timestamp
        List<Time> lowTimeStamp = new ArrayList<>();
        //to store timestamp
        List<Time> highTimeStamp = new ArrayList<>(); 
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

                    //This is meant to use as an index when pulling the TC
                    // names for locating the correct name that is to be paired
                    // with the correct data. If dashNumber is 21 (21 TCs then 
                    // reset) else increment
                    if (dashNumber == 21) {//This block is repeated make a singular function ************************************** For future updates
                        dashNumber = 1;
                    } else {
                        dashNumber++;
                    }
                    //If the TC has a value less than the min required per spec
                    // store that TC name, TC data and time stamp
                    if (cell.getNumericCellValue() <= eMinTemp - 
                            lessTemp/*114.9*/) {

                        flowTcs.add(cell.getNumericCellValue());
                        failedLowTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time lowTime = new Time(tcTime.getTime());
                        lowTimeStamp.add(lowTime);

                    //If the TC has a value greater than the max required per spec
                    // store that TC name, TC data and time stamp
                    } else if (cell.getNumericCellValue() >= eMaxTemp + 
                            lessTemp/*145.1*/) {

                        fhighTcs.add(cell.getNumericCellValue());
                        failedHighTcNames.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellC = row.getCell(locating);
                        Date tcTime = cellC.getDateCellValue();
                        Time highTime = new Time(tcTime.getTime());
                        highTimeStamp.add(highTime);
                    }
                }
            }
            // Jump to the next row in the data file
            startTHoldIndex++;
          // Loop while the starting index hasnt reached the ending index
        } while (startTHoldIndex != endThirdHoldIndex);// outter for loop
        // Method call to Write the descrepant TCs to the report
        writeToReport(flowTcs, fhighTcs, failedLowTcNames, failedHighTcNames, 
                identity, lowTimeStamp, highTimeStamp);
        // Method call to find the ramp rates between the second/third hold
        findTRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    }; // End method complianceThirdHold 
   
   /**
    * Method writeToReport This method take all inputs from the compiance 
    * reporting for tempature and writes them to the report. It takes advantage of a unique 
    * identifier sent from each method call to write the data to the report into 
    * the appropriate location using a switch to call that location
    * @param flowTcs
    * @param fhighTcs
    * @param failedLowTcNames
    * @param failedHighTcNames
    * @param identity
    * @param lowTimeStamp
    * @param highTimeStamp 
    */
   //This uses a switch block which could possibly be made a singular function ************************************** For future updates
   public void writeToReport(List<Double> flowTcs, List<Double> fhighTcs, 
           List<String> failedLowTcNames, List<String> failedHighTcNames, 
           int identity, List<Time> lowTimeStamp, List<Time> highTimeStamp) {
        System.out.println("Write to report");
        int increaseRow;
        int reset = 1;
        int increaseColumn;

        //Try to open the file that contains the report 
        try ( FileInputStream fis = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis)) {
            //Locate the tempature report page
            Sheet sheet = workbook.getSheet("TempOOTs");

            //Write the low and high Tcs to that page 
            switch (identity) {
                
                //First Hold Temp OOTS (Out of Tolerances)
                case 0:
                    increaseRow = 1;
                    increaseColumn = 1;
                    Row row;
                    Cell cellA;
                    int j = 0;
                    for (int i = 0; i < failedLowTcNames.size(); i++) {
                        row = sheet.getRow(increaseRow);

                        //Use j to pull the index value from the associated arrays
                        //Locate the cell at position increase column
                        cellA = row.getCell(increaseColumn);
                        //Save the time stamp at j position
                        String ltStamp = lowTimeStamp.get(j).toString();
                        //Set the timestamp to the increaseColumn position 
                        cellA.setCellValue(ltStamp);

                        //The remaining are similiar, increase the column position
                        //and set the value to the cell 
                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedLowTcNames.get(j));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(flowTcs.get(j));

                        //Now try to write this data to the report 
                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        //Increase row and j values 
                        increaseRow++;
                        j++;
                    }
                // Case 0 for the low temps has no break and this is by design to allow to jump to the next statement 
                // which is to write the highs. The remaining steps follow the same process 
                
                case 10:

                    increaseRow = 1;
                    increaseColumn = 8;
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

                        try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow++;
                        j++;
                    }
                    break;

                //Second Hold Temp OOTS (Out of Tolerances)
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

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                        j++;
                    }

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

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                        j++;
                    }
                    break;

                //Third Hold Temp OOTS (Out of Tolerances)
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

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                        j++;
                    }

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

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
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

        //This portion will write the header data to the report file 
        try ( FileInputStream fis1 = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis1)) {
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
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }; // End method writeToReport 
                            
    /**
     * Method writeRampToReport much like the temp writing method this will 
     * write all of the ramp rates to their position within the report 
     * using the same switch method - For more details review the writing method
     * for temps 
     * @param rampRateList
     * @param identity
     * @param failedTcNames
     * @param timeStamp 
     */
    public void writeRampToReport(List<Double> rampRateList, int identity, 
            List<String> failedTcNames, List<Time> timeStamp) {

        try ( FileInputStream fis = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis)) {

            int increaseRow;
            int reset = 1;
            int increaseColumn;
            Sheet sheet = workbook.getSheet("RampRateOOTs");

            switch (identity) {
                case 3:
                    increaseRow = 1;
                    increaseColumn = 1;

                    Row row;
                    Cell cellA;

                    for (int i = 0; i < failedTcNames.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString();
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedTcNames.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(rampRateList.get(i));

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;

                case 4:
                    increaseRow = 1;
                    increaseColumn = 8;

                    for (int i = 0; i < failedTcNames.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString();
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedTcNames.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(rampRateList.get(i));

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;

                case 5:
                    increaseRow = 1;
                    increaseColumn = 15;

                    for (int i = 0; i < failedTcNames.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString();
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedTcNames.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(rampRateList.get(i));

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;

                case 6:
                    increaseRow = 1;
                    increaseColumn = 22;

                    for (int i = 0; i < failedTcNames.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString();
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedTcNames.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(rampRateList.get(i));

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }; //End method writeRampToReport 

    /**
     * Method findRampRate this method will search by column comparing the 
     * value in a row and the next rows data which subtracts and compares to the 
     * required ramp rate value per spec. Once it reviews an entire column it moves
     * to the next column until the end of the TC data is found by way of array end
     * this is different than the temp because it reviews by column where the temp
     * reviews by row 
     * ******************************************In reality this entire sequence needs to be a singular function shared by all three ramps - First Priority
     * @param curespec
     * @param rowIndex
     * @param columnIndex
     * @param dataFilePath
     * @param holdIndex
     * @param endHoldIndex
     * @param sheet 
     */
   public void findRampRate(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, int holdIndex,
            int endHoldIndex, Sheet sheet) {
        
        double currentNum = 0;
        double previousNum = 0;
        int identity = 3;
        int j = 0;
        int startingPoint = rowIndex + 2;
        int nextCol = startingPoint;
        int dashNumber = 0;
        rampRateMax = curespec.getRampRate();
        List<String> failedRampTcs = new ArrayList<>();
        List<Double> columnDeltas = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        int locating = 0;

        System.out.println("Find Ramp");
        
        for (int item1 : columnIndex) {
            do {
                Row row = sheet.getRow(holdIndex);
                if (row != null) {
                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    //This deals with the issue where th value from the last position of any column
                    //Is compared to the first value of the new column which results in a large failed ramp
                    //This corrects it by setting the new first value (new column) to itself as it doesnt 
                    //need to be reviewed in the same way
                    if (holdIndex == 25) {
                        previousNum = currentNum;
                    }
                    if (dashNumber == 21) {
                        dashNumber = 1;
                    }

                    //Assign the ramp to rampRate taking the absolute non negative value
                    double rampRate = Math.abs(currentNum - previousNum);

                    if (rampRate >= rampRateMax) {
                        columnDeltas.add(rampRate);
                        failedRampTcs.add(tcNames.get(dashNumber));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);
                    }
                }
                
                previousNum = currentNum;
                holdIndex++;
            } while (holdIndex != endHoldIndex);
            dashNumber++;
            //Next column is set to the holding starting point 
            holdIndex = nextCol;
        }
        //Method call to write the ramps to the report 
        writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);
        
    }; //End method findRampRate

   /**
    * Method findSRampRate this method will search by column comparing the 
    * value in a row and the next rows data which subtracts and compares to the 
    * required ramp rate value per spec. Once it reviews an entire column it moves
    * to the next column until the end of the TC data is found by way of array end
    * this is different than the temp because it reviews by column where the temp
    * reviews by row 
    * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
   public void findSRampRate(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
        System.out.println("Find SRamp");
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

        //While there is an item in the columnIndex array do this
        for (int item1 : columnIndex) {
            do {
                Row row = sheet.getRow(startingPoint);
                if (row != null) {
                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();
                    if (startingPoint == nextCol) {
                        previousNum = currentNum;
                    }
                    if (dashNumber == 21) {
                        dashNumber = 1;
                    }

                    //Assign the ramp to rampRate taking the absolute non negative value
                    double rampRate = Math.abs(currentNum - previousNum);

                    //If value exceeds the ramp rate per spec add it to the failed array
                    if (rampRate >= rampRateMax) {
                        columnDeltas.add(rampRate);
                        failedRampTcs.add(tcNames.get(dashNumber));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);
                    }
                }
                previousNum = currentNum;
                startingPoint++;

            } while (startingPoint != endingPoint);
            startingPoint = nextCol;
            dashNumber++;
        }
        //Method call to write the ramp rates to the report 
        writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);
        
    }; // End method findSRamp 

   /**
    * Method findTRampRate: this method will search by column comparing the 
    * value in a row and the next rows data which subtracts and compares to the 
    * required ramp rate value per spec. Once it reviews an entire column it moves
    * to the next column until the end of the TC data is found by way of array end
    * this is different than the temp because it reviews by column where the temp
    * reviews by row 
    * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
   public void findTRampRate(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
        System.out.println("Find TRamp");
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
        List<Double> columnDeltas = new ArrayList<>();
        List<String> failedRampTcs = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();

        //While there is an item in the columnIndex array do this
        for (int item1 : columnIndex) {

            do {
                Row row = sheet.getRow(startingPoint);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();

                    if (startingPoint == nextCol) {
                        previousNum = currentNum;
                    }
                    if (dashNumber == 21) {
                        dashNumber = 1;
                    }

                    //Assign the ramp to rampRate taking the absolute non negative value
                    double rampRate = Math.abs(currentNum - previousNum);
                    
                    //If value exceeds the ramp rate per spec add it to the failed array
                    if (rampRate >= rampRateMax) {
                        columnDeltas.add(rampRate);
                        failedRampTcs.add(tcNames.get(dashNumber));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);
                    }

                }
                previousNum = currentNum;
                startingPoint++;

            } while (startingPoint != endingPoint);
            startingPoint = nextCol;
            dashNumber++;

        }
        //Method call to write the ramp rates to the report
        writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);
        //Method call to find the last ramp rate that focuses on the cool down step
        findCoolRampRate(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    }; // End method findTRampRate
    
   /**
    * Method findCoolRampRate: this method will search by column comparing the 
    * value in a row and the next rows data which subtracts and compares to the 
    * required ramp rate value per spec. Once it reviews an entire column it moves
    * to the next column until the end of the TC data is found by way of array end
    * this is different than the temp because it reviews by column where the temp
    * reviews by row 
    * ******************************************In reality this entire sequence needs to be a singular function shared by all three holds - First Priority
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
   public void findCoolRampRate(CureSpec curespec, int rowIndex, List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
        System.out.println("Find cool ramp");
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

        //While there is an item in the columnIndex array do this
        for (int item1 : columnIndex) {
            do {
                Row row = sheet.getRow(startingPoint);
                if (row != null) {

                    Cell cell = row.getCell(item1);
                    currentNum = cell.getNumericCellValue();

                    if (startingPoint == nextCol) {
                        previousNum = currentNum;
                    }
                    if (dashNumber == 21) {
                        dashNumber = 1;
                    }

                    //Assign the ramp to rampRate taking the absolute non negative value
                    double rampRate = Math.abs(previousNum - currentNum);
                    
                    //If value exceeds the ramp rate per spec add it to the failed array
                    if (rampRate >= rampRateMax) {
                        columnDeltas.add(rampRate);
                        failedRampTcs.add(tcNames.get(dashNumber));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);
                    }

                }
                //If the next row doesnt have data then end of data found set the ending point equal to the starting point to 
                // break from the while loop 
                if (row == null) {
                    endingPoint = startingPoint;
                    break;
                }
                previousNum = currentNum;
                startingPoint++;
                endingPoint++;

            } while (startingPoint != endingPoint);
            //Move to the next column
            startingPoint = nextCol;
            dashNumber++;
            dataEnd = endingPoint;

        } // End for loop 

        //Method call to write failed ramps to the report 
        writeRampToReport(columnDeltas, identity, failedRampTcs, timeStamp);
        
        //Method call to review the vacuum TC data 
        findVacuum(curespec, rowIndex, columnIndex, dataFilePath, sheet);
    }; //End method findCoolRampRate

    
   /**
    * Method findVacuum: This will look for the vacuum TCs by pulling the indexes
    * in the vacuumIndex array and using the data at those positions within the data
    * file to review per spec the vacuum requirements 
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
   public void findVacuum(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        
        int dashNumber = 0;
        int identity = 7;
        //Get the vacuum req from the cure spec
        vacPressure = curespec.getVacuum();
        //The requirements have changed since beginning 
        //This value will need recalculated before implementation 
        minVacPressure = vacPressure - 6;
        List<String> failedVacInHg = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        int locating = 0;

        do {
            Row row = sheet.getRow(dataStart);
            if (row != null) {
                for (int item : vacuumIndex) {
                    Cell cell = row.getCell(item);

                    if (dashNumber == 9) { //Make into a var that stores the number of TCs found instead of providing this value
                        dashNumber = 0;
                    }

                    //Assign the value of the cell in the absolute non negative value to vacPres
                    double vacPres = Math.abs(cell.getNumericCellValue());

                    //If the vac pressure doesnt meet the requirement per spec add to the failedVac array
                    if (vacPres <= minVacPressure) {/*23.9*/

                        vacInHg.add(vacPres);
                        failedVacInHg.add(vacNames.get(dashNumber));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);
                    }
                    dashNumber++;
                }
            }
            dataStart++;

        } while (dataStart != dataEnd);
        dataStart = 0;
        //Method call for writing the descrepant vacuum values to the report 
        writeVacuumToReport(vacInHg, identity, failedVacInHg, timeStamp);
        System.out.println("Writing Vaccum Completed");
        
        //Method call to Verify the 9002 requirements 
        find9002(curespec, rowIndex, columnIndex, dataFilePath, sheet);

    }; //End method findVacuum 
 
   /**
    * Method writeVacuumToReport This method takes the vacuum data identified 
    * as failing and writes it to the report 
    * @param vacInHg
    * @param identity
    * @param failedVacInHg
    * @param timeStamp 
    */
   public void writeVacuumToReport(List<Double> vacInHg, int identity, 
           List<String> failedVacInHg, List<Time> timeStamp) {
        System.out.println("Writing Vaccum");
        try ( FileInputStream fis = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis)) {

            int increaseRow;
            int reset = 1;
            int increaseColumn;
            //Locate the sheet with the following name 
            Sheet sheet = workbook.getSheet("VacuumOOTs");

            //Do the following if the required ID is recieved
            switch (identity) {
                case 7:
                    increaseRow = 1;
                    increaseColumn = 1;

                    Row row;
                    Cell cellA;

                    for (int i = 0; i < failedVacInHg.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        //Get the cell at increaseColumn location in sheet and add the timestamp 
                        //and the rest of the data (tc name, and data)
                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString(); 
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failedVacInHg.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(vacInHg.get(i));

                        try ( FileOutputStream fos = new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Vacuum Write Complete");
    }; //End method writeVacuumToReport 

   /**
    * Method find9002: This method observes the temps throughout the entire cure
    * for breaks of the max tempature requirement per spec. If it breaks it stores those
    * TCs to add to the report 
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
   public void find9002(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, Sheet sheet) {
        System.out.println("Find 9002");
        int identity = 8;
        int dashNumber = 0;
        //Grab the requirement in the spec
        temp9002 = curespec.getTemp9002();
        List<Double> nineTwoHighTemp = new ArrayList<>();
        List<String> failednineTwoHighTemp = new ArrayList<>();
        List<Time> timeStamp = new ArrayList<>();
        int locating = 0;

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
                    
                    //If requirement fails add to the array 
                    if (cell.getNumericCellValue() >= temp9002/*360.1*/) {
                        nineTwoHighTemp.add(cell.getNumericCellValue());
                        failednineTwoHighTemp.add(tcNames.get(dashNumber - 1));
                        //This block is repeated make a singular function ************************************** For future updates
                        Cell cellB = row.getCell(locating);
                        Date tcTime = cellB.getDateCellValue();
                        Time time = new Time(tcTime.getTime());
                        timeStamp.add(time);

                    }
                }

            }
            dataStart2++;
        } while (dataStart2 != dataEnd);

        //Method call to write to the report 
        write9002ToReport(nineTwoHighTemp, identity, failednineTwoHighTemp, timeStamp);

    }; //End method find9002
 
   /**
    * Method write9002ToReport: This will take whatever was found as noncompliant 
    * and write it in the appropriate section within the report 
    * @param nineTwoHighTemp
    * @param identity
    * @param failednineTwoHighTemp
    * @param timeStamp 
    */
   public void write9002ToReport(List<Double> nineTwoHighTemp, 
           int identity, List<String> failednineTwoHighTemp, 
           List<Time> timeStamp) {
        System.out.println("Write 9002");
        try ( FileInputStream fis = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis)) {

            int increaseRow;
            int reset = 1;
            int increaseColumn;
            //Open the sheet with the following name 
            Sheet sheet = workbook.getSheet("9002OOTs");

            //If unique ID is provided do the following
            switch (identity) {
                case 8:
                    increaseRow = 1;
                    increaseColumn = 1;

                    Row row;
                    Cell cellA;

                    for (int i = 0; i < failednineTwoHighTemp.size(); i++) {

                        row = sheet.getRow(increaseRow);

                        //Start at determined column (increaseColumn) and add the timestamp
                        //Increase column and add the TC name and then the data itself
                        cellA = row.getCell(increaseColumn);
                        String tStamp = timeStamp.get(i).toString();
                        cellA.setCellValue(tStamp);

                        cellA = row.getCell(increaseColumn + 1);
                        cellA.setCellValue(failednineTwoHighTemp.get(i));

                        cellA = row.getCell(increaseColumn + 2);
                        cellA.setCellValue(nineTwoHighTemp.get(i));

                        try ( FileOutputStream fos = 
                                new FileOutputStream(reportFilePath)) {
                            workbook.write(fos);
                        } catch (IOException e) {
                            e.printStackTrace();
                        };
                        increaseRow = increaseRow + 1;
                    }
                    break;

            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        //This try catch block will locate the report file and open it once the program has 
        // completed the review process 
        try ( FileInputStream fis = new FileInputStream(reportFilePath);  
                Workbook workbook = new XSSFWorkbook(fis)) {
            Desktop.getDesktop().open(new File(reportFilePath));
        } catch (IOException e) {
            e.printStackTrace();
        }
    };// End method write9002ToReport 

 
   /**
    * Method getTimeStamp this is a future effort to isolate the timestamps into a single method but is not yet implemented 
    * @param curespec
    * @param rowIndex
    * @param columnIndex
    * @param dataFilePath
    * @param sheet 
    */
 /*  public void getTimeStamp(CureSpec curespec, int rowIndex, 
           List<Integer> columnIndex, String dataFilePath, Sheet sheet) {

        int locating = 0;
        boolean condition = true;
        int newRowIndex = rowIndex + 2;

        List<Time> timeStamp = new ArrayList<>();
        sheet = openExcel(dataFilePath);

        do {
            Row row = sheet.getRow(newRowIndex);
            if (row != null) {
                Cell cell = row.getCell(locating);
                Date tcTime = cell.getDateCellValue();
                Time time = new Time(tcTime.getTime());
                timeStamp.add(time);
                System.out.println(time);
            } else {
                condition = false;
            }
            newRowIndex++;

        } while (condition = true);
    };//End method getTimeStamp
*/
} //End Subclass ExcelHandler

