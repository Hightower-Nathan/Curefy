package CurefyPkg;
/**
* @Author Name: Nathan Hightower
* @Project Name: Curefy
* @Date: Feb 4, 2025
* @Subclass Curespec Description: This will handle establishing cure spec's
*/ 

//Imports
import java.util.List;
import java.util.ArrayList;

//Begin Subclass CureSpec
public class CureSpec {
    
    protected double vacuum;
    protected double aMaxTemp, cMaxTemp, eMaxTemp; 
    protected double aMinTemp, cMinTemp, eMinTemp;
    protected double rampRate; 
    protected double rampTime;
    protected double bHoldTime, dHoldTime, fHoldTime;
    protected double vacuumDecay;
    protected double coolTemp; 
    protected String tcName; //May need to be an arraylist<> 
    protected double delta;
    protected double lessTemp = 0.1;
    protected double temp9002 = 360.1;
    
    //In the furture use the constructor to initilize variables.
    //This will allow the appropriate spec to be created by the number of 
    //Args that are passed to it. 
    
    //Constructors 
    public CureSpec(){};//This is for the ExcelHandler to extend. I need to work on getting everything organized 
    
    public CureSpec(List<Double> list){
        System.out.println("CureSpec Created!");
        setCureValues(list);
    }
    
    /**
     * Method: Sets the values for tc names - may change to arraylist<>
     * @param tcNames 
     */
    public void setTcNames(String tcNames){
        tcName = tcNames;
    }//End method
    
    //This is a test to set all required values at the same time 
    
    /**
     * Method: Sets all the values from the DB at once...
     * Has to be a better way of doing this though...
     * Look into!!!!!!!!!!!
     * I do not like hard coding values.
     * @param list 
     */
    public void setCureValues(List<Double> list){
    
        vacuum = list.get(0);
        aMaxTemp = list.get(1);
        aMinTemp = list.get(2);
        rampRate = list.get(3);
        rampTime = list.get(4);
        bHoldTime = list.get(5);
        cMaxTemp = list.get(6);
        cMinTemp = list.get(7);
        dHoldTime = list.get(8); 
        eMaxTemp = list.get(9);
        eMinTemp = list.get(10);
        vacuumDecay = list.get(11);
        fHoldTime = list.get(12);
        coolTemp = list.get(13);
        delta = list.get(14);
       // printVars();
  
    };//End method
    
      //This is to verify the sets are correct..they are 
    public void printVars(){
        System.out.println(vacuum);
        System.out.println(aMaxTemp);
        System.out.println(aMinTemp);
        System.out.println(rampRate);
        System.out.println(rampTime);
        System.out.println(bHoldTime);
        System.out.println(cMaxTemp);
        System.out.println(cMinTemp);
        System.out.println(dHoldTime);
        System.out.println(eMaxTemp);
        System.out.println(eMinTemp);
        System.out.println(vacuumDecay);
        System.out.println(fHoldTime);
        System.out.println(coolTemp);
        System.out.println(delta);
    }

    ////////////////////////////////////////////////////////////////////////////
    //May not need get methods since the verify classes will implement CureSpec 
    //all variables should be able to be accessed if not private then protected
    ////////////////////////////////////////////////////////////////////////////
    
    /**
     * Method: Returns the value stored in temp9002
     * @return 
     */
    public double getTemp9002(){
        return temp9002;
    }// End method 
    
    
    /**
     * Method: Returns the value stored in lessTemp
     * @return 
     */
    public double getLessTemp(){
        return lessTemp;
    }//End Method
    
    /**
     * Method: Returns the value stored in vacuum
     * @return 
     */
    public double getVacuum(){
        return vacuum; 
    }//End method 
    
    /**
     * Method: Returns the value stored in aMaxTemp
     * @return 
     */
    public double getabMaxTemp(){
        return aMaxTemp;
    }//End method
    
    /**
     * Method: Returns the value stored in cMaxTemp
     * @return 
     */
    public double getcdMaxTemp(){
        return cMaxTemp;
    }//End method
    
    /**
     * Method: Returns the value stored in eMaxTemp
     * @return 
     */
    public double getefMaxTemp(){
        return eMaxTemp;
    }//End method
    
    /**
     * Method: Returns the value of aMinTemp
     * @return 
     */
    public double getabMinTemp(){
        return aMinTemp;
    }// End method
    
    /**
     * Method: Returns the value stored in cMinTemp
     * @return 
     */
    public double getcdMinTemp(){
        return cMinTemp;
    }//End method
    
    /**
     * Method: Returns the value stored in eMinTemp
     * @return 
     */
    public double getefMinTemp(){
        return eMinTemp;
    }//End method
    
    /**
     * Method: Returns the value stored in rampRate
     * @return 
     */
    public double getRampRate(){
        return rampRate;
    }//End method
    
    /**
     * Method: Returns the value stored in rampTime
     * @return 
     */
    public double getRampTime(){
        return rampTime;
    }//End method
    
    /**
     * Method: Returns the value stored bHoldTime
     * @return 
     */
    public double getbHoldTime(){
        return bHoldTime;
    }//End method
    
    /**
     * Method: Returns the value stored in dHoldTime
     * @return 
     */
    public double getdHoldTime(){
        return dHoldTime;
    }//End method
    
    /**
     * Method: Returns the value stored in fHoldTime
     * @return 
     */
    public double getfHoldTime(){
        return fHoldTime;
    }//End method
    
    /**
     * Method: Returns the value stored in vacuumDecay
     * @return 
     */
    public double getVacuumDecay(){
        return vacuumDecay; 
    }//End method 
    
    /**
     * Method: Returns the value stored in coolTemp
     * @return 
     */
    public double getCoolTemp(){
        return coolTemp;
    }//End method 
    
    /**
     * Method: Returns the values stored in tcName
     * May change to arraylist<>
     * @return 
     */
    public String getTcName(){
        return tcName;
    }//End method
    
    /**
     * Method: Returns the value stored in delta
     * @return 
     */
    public double getDelta(){
        return delta;
    }//End method 
    
} //End Subclass CureSpec
