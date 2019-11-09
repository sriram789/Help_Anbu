package HelpAnbu;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Scanner;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class createNewExcel  
{  
	public static Scanner scanner = new Scanner(System.in);
	
	public static void main(String[] args)   
	{  
		System.out.println("Enter the Input Location");
		String location = scanner.next();
		System.out.println("Enter the Output Location");
		String outputLocation = scanner.next();
		int count;
		HashMap<String, Integer> Distribution_set = new HashMap<>(); 
		try  
		{  
		File file = new File(location);   					//creating a new file instance  
		FileInputStream fis = new FileInputStream(file);   	//obtaining bytes from the file    
		XSSFWorkbook wb = new XSSFWorkbook(fis);   			//creating Workbook instance that refers to .xlsx file
		XSSFSheet sheet = wb.getSheetAt(0);     			//creating a Sheet object to retrieve object  
		Iterator<Row> itr = sheet.iterator();    			//iterating over excel file  
		itr.next();
		while (itr.hasNext())                 
		{  
		Row row = itr.next();  
		Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column 
		cellIterator.next();
		
		
		Cell key = cellIterator.next();
		
		if(Distribution_set.containsKey(key.getStringCellValue())) {
			count = Distribution_set.get(key.getStringCellValue())+1;
			Distribution_set.put(key.getStringCellValue(), count);
		}
		else {
			Distribution_set.put(key.getStringCellValue(), 0);
		}	
		}  
		
		wb.close();
		}  
		catch(Exception e)  
		{  
		e.printStackTrace();  
		}
		
		//Writing data to the Excel
		
		// Blank workbook 
        XSSFWorkbook workbook = new XSSFWorkbook(); 
  
        // Create a blank sheet 
        XSSFSheet sheet = workbook.createSheet("Description Count");
        
		Set<String> keyset = Distribution_set.keySet(); 
		
        int rownum = 0; 
        
        for (String key : keyset) { 
            // this creates a new row in the sheet 
            Row row = sheet.createRow(rownum++); 
                // this line creates a cell in the next column of that row 
                Cell cell0 = row.createCell(0); 
                Cell cell1 = row.createCell(1);  
                cell0.setCellValue((String)key); 
                cell1.setCellValue((Integer)Distribution_set.get(key)); 
        } 
        try { 
            // this Writes the workbook gfgcontribute 
            FileOutputStream out = new FileOutputStream(new File(outputLocation)); 
            workbook.write(out); 
            out.close(); 
            System.out.println(outputLocation+" written successfully on disk."); 
            workbook.close();
        } 
        catch (Exception e) { 
            e.printStackTrace(); 
        } 
	}    
 
} 