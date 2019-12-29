package helpAnbu;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.Scanner;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.*;  
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
public class createNewExcel  
{  
	public static Scanner scanner = new Scanner(System.in);
	static Set<String> Login = new LinkedHashSet<String>(); 
	static Set<String> Registration = new LinkedHashSet<String>(); 
	static Set<String> SSO = new LinkedHashSet<String>();
	static Set<String> PCP = new LinkedHashSet<String>(); 
	static Set<String> MyCoverage = new LinkedHashSet<String>(); 
	static Set<String> ClaimCentre = new LinkedHashSet<String>(); 
	static Set<String> EOB = new LinkedHashSet<String>();
	static Set<String> HSA = new LinkedHashSet<String>(); 
	static Set<String> HCA = new LinkedHashSet<String>(); 
	static Set<String> DoctorAndHospitals = new LinkedHashSet<String>(); 
	static Set<String> DynatraceAlert = new LinkedHashSet<String>(); 
	
	static ArrayList<Set<String>> groupNames = new ArrayList<Set<String>>();
	
	public static String getSetName(ArrayList<Set<String>> array, String data) {
		String returnName = "NULL";
		for(Set<String> set:array) {
			if(set.contains(data)) {
				returnName = set.stream().findFirst().get();
				break;
			}
		}
		return returnName;
	}
	public static void main(String[] args)   
	{  
		System.out.println("Enter the Input Location");
		String location = scanner.next();
		System.out.println("Enter the Output Location");
		String outputLocation = scanner.next();
		int count;
		
		
		Login.add("Login");
		Login.add("password Issue");
		Login.add("BAM CSR");
		SSO.add("SSO");
		SSO.add("prescription");
		SSO.add("prime therapautics");
		Registration.add("Registration");
		PCP.add("PCP");
		MyCoverage.add("MyCoverage");
		ClaimCentre.add("BAM Claim Centre");
		EOB.add("EOB");
		HSA.add("HSA");
		HCA.add("HCA");
		DoctorAndHospitals.add("Provider Finder");
		DynatraceAlert.add("Alert");
		
		
		groupNames.add(ClaimCentre);
		groupNames.add(DynatraceAlert);
		groupNames.add(EOB);
		groupNames.add(MyCoverage);
		groupNames.add(Registration);
		groupNames.add(DoctorAndHospitals);
		groupNames.add(HCA);
		groupNames.add(HSA);
		groupNames.add(Login);
		groupNames.add(PCP);
		groupNames.add(SSO);
		
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
		
		if(Distribution_set.containsKey(getSetName(groupNames,key.getStringCellValue()))) {
			count = Distribution_set.get(getSetName(groupNames,key.getStringCellValue()))+1;
			Distribution_set.put(getSetName(groupNames,key.getStringCellValue()), count);
		}
		else {
			Distribution_set.put(getSetName(groupNames,key.getStringCellValue()), 1);
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