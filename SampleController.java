
package CurefyPkg;
/**
* @Course: SDEV 250 ~ Java Programming I
* @Author Name: Nathan
* @Assignment Name: CurefyPkg
* @Date: Apr 9, 2025
* @Subclass SampleController Description: Controller for the UI to begin the actual reviewing process 
*/ 
//Imports
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.ProgressBar;
import javafx.stage.FileChooser;
import javafx.scene.control.Button;
import javafx.stage.Stage;
import java.io.File;
import javafx.scene.layout.VBox;
import javafx.scene.Scene;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.concurrent.Task;

 
//Added these to test
import java.io.FileInputStream;
import java.io.IOException; 
import java.util.ArrayList;
import java.util.List;

//Begin Subclass SampleController
public class SampleController extends ExcelHandler  { 
    
   String excelFilePath;// = "C:/Program Files/NetBeans-12.6/WorkSpace/Cure_Specs.xlsx";
   String dataFilePath;// = "C:/Program Files/NetBeans-12.6/WorkSpace/Test_Data_Curefy(1).xlsx";
   int columnIndex = 1; //Passes to the excelHandler.readColumn method 
   
 
    @FXML
    private void handleButtonClick() {
   
        
        Alert alert = new Alert(AlertType.INFORMATION);
        alert.setTitle("Information Dialog");
        alert.setHeaderText(null);
        alert.setContentText("Confirm Run?");
        alert.showAndWait();
   

   //create a new task to run 
   Task<Void> cureTask = new Task<Void>(){

            @Override
            protected Void call() throws Exception {
                
//throw new UnsupportedOperationException("Not supported yet."); // Generated from nbfs://nbhost/SystemFileSystem/Templates/Classes/Code/GeneratedMethodBody
  
   //These are all of the methods that begin the cure program they chain together running, returning and providing 
   //data for the next call until the entire cure is reviewed. 
   //They are called by the UI and Thread that has been assigned to run them via call method 
   ExcelHandler excelhandler = new ExcelHandler();
   List<Double> columnValues = 
           excelhandler.readColumn(excelFilePath, columnIndex);
   
   CureSpec curespec = new CureSpec(columnValues);
   
   ExcelHandler ex1 = new ExcelHandler(curespec); 
   
   int returnedRowIndex = 
           excelhandler.findRowIndex(dataFilePath);
   
   List<String> tcNames = 
           excelhandler.readTcNames(dataFilePath, returnedRowIndex);
   
   List<Integer> returnedColumnIndex = 
           excelhandler.findColumnIndex(dataFilePath, returnedRowIndex);
   
   excelhandler.findFirstHold(curespec, returnedRowIndex, 
           returnedColumnIndex, dataFilePath);
   
   System.out.println("All methods called");
   
   //Once the program is done running update the progress bar until at 100 percent 
   for (int i = 0; i <= 100; i++){
       updateProgress(i, 100);
       Thread.sleep(50);
       
   }
            return null;
            }
        
            
   };//end cureTask
   
        //Bind the progress bar to the running task so that it shows feedback the program is running 
        myProgressBar.progressProperty().bind(cureTask.progressProperty());
        
         //Start the task in a new thread
        new Thread(cureTask).start(); 
    }
    
 
    
    //Create a new label in the UI to display which file was selected or revelvant message pertaining to the file type selected 
    @FXML
    private Label mySelectedFile = new Label();
    
    //Filechooser for opening an importing the data file needing reviewed 
    @FXML
    private void handleButtonClick1() { //Stage primaryStage as var
        
        String extension = ".xlsx";
        //Set the stage 
        Stage primaryStage = new Stage();
        //Create fileChooser 
        FileChooser fileChooser = new FileChooser();
        //Set the filechooser default directory 
        fileChooser.setInitialDirectory(new File(
                "C:/Program Files/NetBeans-12.6/WorkSpace"));
        //Set title 
        fileChooser.setTitle("Open Data File");

            //Display the file chooser 
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            
            if (selectedFile != null){
            //added this stuff for basic file type validation but could be done better........................Future update and priority 
                String fileType = selectedFile.getName();
                if(fileType.endsWith(extension)){
                    System.out.println("User selected XLSX file");
                    dataFilePath = selectedFile.getAbsolutePath();
                    mySelectedFile.setText(selectedFile.getAbsolutePath());
                }
                else {
                    System.out.println("Wrong File Type Selected");
                    mySelectedFile.setText("Wrong File Type Selected.."
                            + "File must be '.xlsx'.. Try Again");
                }
            }
            else{
                System.out.println("Invalid Selection");
            }
    }
    
    //Create a combo box for the spec selection 
    @FXML
    private ComboBox<String> myComboBox;
    
    //Create a progress bar for providing feedback to the user 
    @FXML
    private ProgressBar myProgressBar;

    //Add elements to the combo box 
    @FXML
    private void initialize() { 
        //Add spec to combo box 
        myComboBox.getItems().addAll("ACS-PRS-5999.102");
        final String[] chosenSpec = new String[1];
        //Select spec 
        myComboBox.setOnAction(event -> {
            chosenSpec[0] = myComboBox.getSelectionModel().getSelectedItem();
            System.out.println("Selected: " + chosenSpec[0]);
            
            //Add selected spec as a string for confirmation 
            String selection = chosenSpec[0].toString();
            
            //if string equals this do the following 
            if(selection == "ACS-PRS-5999.102"){
                
                excelFilePath = "C:/Program Files/NetBeans-12.6/WorkSpace/Cure_Specs.xlsx";
                System.out.println("5999.102 was selected and file has been uploaded");
 
            }
            else{
                System.out.println("File not uploaded");
            }
        });

                }

     
  
  
     
} //End Subclass SampleController
 