package CurefyPkg;

/**
 * @Author Name: Nathan Hightower
 * @Project Name: Curefy
 * @Date: Jan 25, 2025 - April 19, 2025
 * @Description: This is a Capstone project meant to examine autoclave cure data
 * and generate reports
 */

//Imports
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import javafx.scene.layout.StackPane;
import javafx.scene.paint.Color;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.Label;
import javafx.concurrent.Task;

//Begin Class Curefy
public class Curefy extends Application {

    /**
     * Method start sets the UI stage and displays it for use
     *
     * @param primaryStage
     * @throws Exception
     */
    @Override
    public void start(Stage primaryStage) throws Exception {
        Parent root = FXMLLoader.load(getClass().getResource("sample.fxml"));
        primaryStage.setTitle("Curefy - Autoclave Cure Verification");
        Scene scene = new Scene(root, 680, 275);
        primaryStage.setScene(scene);
        primaryStage.show();

    }// End method start

//Begin Main Method  
    public static void main(String[] args) {

        //Begin the application 
        launch(args);

    } //End Main Method

} //End Class Curefy
         