

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class RunnerClass {
	
	
  @Test
  public void ReadExcelFile(String filePath, String fileName, String sheetName ) throws IOException {
	//Create an object of File class to open xlsx file
	  File file=new File(filePath+"\\"+fileName);
	  
	//Create an object of FileInputStream class to read excel file
	  FileInputStream inputstream = new FileInputStream(file);
	  
	  Workbook workbook=null;
	  
	//Find the file extension by splitting file name in substring  and getting only extension name
	  String fileExtensionName = fileName.substring(fileName.indexOf("."));

	    //Check condition if the file is xlsx file

	    if(fileExtensionName.equals(".xlsx")){ 
	    	//If it is xlsx file then create object of XSSFWorkbook class	  
	    	workbook = new XSSFWorkbook(inputstream);
	    }
	  //Check condition if the file is xls file
	    else if(fileExtensionName.equals(".xls"))
	    {
	    	 //If it is xls file then create object of HSSFWorkbook class
	    	workbook = new HSSFWorkbook(inputstream);
	    	
	    }
	  //Read sheet inside the workbook by its name
	    Sheet TestPop=workbook.getSheet(sheetName);
	    
	    //Find number of rows in excel file
	    int rowCount=TestPop.getLastRowNum()-TestPop.getFirstRowNum();
	    
	  //Create a loop over all the rows of excel file to read it
	    for(int i=0; i<rowCount;i++){
	    	Row row= TestPop.getRow(i);
	    	
	   //Create a loop to print cell values in a row
	    	for(int j=0 ; j<row.getLastCellNum();j++){
	    		//Print Excel data in console
	    		System.out.println(row.getCell(j).getStringCellValue()+"||");
	    	}
	    	System.out.println();
	    	
	    }
  } 
  @Test()
  public void passingvalue() throws IOException{
	//Create an object of ReadExcelFile class
	  RunnerClass rd=new RunnerClass();
	  
	//Prepare the path of excel file
	  String filePath =System.getProperty("user.dir")+"//FileOperations//TestData.xlsx";
	  
	//Call read file method of the class to read data
	  rd.ReadExcelFile(filePath, "TestData.xlsx", "TestPop");
	  
  }
}

