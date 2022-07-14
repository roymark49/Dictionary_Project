package excel_File_Read;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Read_exl_file {
	
	static ArrayList<String> myArray;

	public static void main(String[] args) throws Exception {
		
		System.out.print("Enter the File Path: ");
		
		Scanner userinput = new Scanner(System.in);
		String path =userinput.nextLine();
		
		doesFileExist(path);
	}
	
	public static ArrayList<String> readExcelFile() throws Exception {
		
		try {
			InputStream input = new FileInputStream("data\\DictionaryFile.xlsx");
			
			Workbook wb = WorkbookFactory.create(input);
			Sheet ws = wb.getSheet("Dictionary");
			
			int numberOfRows= ws.getPhysicalNumberOfRows();
			
			myArray = new ArrayList<String>();
			
			for(int i=0; i<numberOfRows ; i++ ) {
				myArray.add(ws.getRow(i).getCell(0).toString());
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
		return myArray;
		
	}
	
	public static void doesFileExist(String path) throws Exception {
		
		if(path.contains("data\\DictionaryFile.xlsx")) {
			System.out.println("File found at the path");
			System.out.println("Word 1: " + readExcelFile().get(0));
			System.out.println("Meaning 1: " + readExcelFile().get(1));
			System.out.println("Meaning 2: " + readExcelFile().get(2));
			
			System.out.println("Word 2: " + readExcelFile().get(3));
			System.out.println("Meaning 1: " + readExcelFile().get(4));
			System.out.println("Meaning 2: " + readExcelFile().get(5));
			
			System.out.println("Word 3: " + readExcelFile().get(6));
			System.out.println("Meaning 1: " + readExcelFile().get(7));
			
		} else {
			System.out.println("file not found at the path");
		}
	}
	
}