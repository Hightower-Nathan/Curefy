package CurefyPkg;
/** 
* @Course: SDEV 250 ~ Java Programming I
* @Author Name: Nathan
* @Assignment Name: testapachepoi
* @Date: Jan 28, 2025
* @Subclass MyRunnable Description:
*/
//Imports
//Begin Subclass MyRunnable
public abstract class MyRunnable implements Runnable{
    
    public void run() {
        try{
        System.out.println("Thread" + Thread.currentThread().getId() + " is running");
        }
        catch (Exception e) {
            System.out.println("Exception is caught");
        }
    }
} //End Subclass MyRunnable
