package com.pits.worksheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorksheetMain {

	public static void main(String[] args) throws IOException {

		
		
		String exportSheetName = null;
		String customerSheetName = null;
		String addressSheetName = null;
		
		Scanner sc = new Scanner(System.in);
		
		System.out.println("Enter Export Sheet Name with extension:");
		exportSheetName = sc.nextLine();
		
		System.out.println("Enter Customer Sheet name with extension:");
		customerSheetName = sc.nextLine();
		
		System.out.println("Enter Address Sheet name with Extension:");
		addressSheetName = sc.nextLine();
		
		long start = System.currentTimeMillis();
		
		FileInputStream exportfile = new FileInputStream(
				new File(".\\source\\"+exportSheetName));

		XSSFWorkbook workbookExport = new XSSFWorkbook(exportfile);
		XSSFSheet exportSheet = workbookExport.getSheet("Export");

		int row = exportSheet.getLastRowNum();

		HashMap<String, String> keyValuePair = new HashMap<String, String>();

		for (int r = 0; r <= row; r++) {
			String key = exportSheet.getRow(r).getCell(0).getStringCellValue();
			String value = exportSheet.getRow(r).getCell(1).getStringCellValue();
			keyValuePair.put(key, value);
		}

		/*
		 * for(Map.Entry entry:data.entrySet()) {
		 * System.out.println(entry.getKey()+" "+entry.getValue()); }
		 */

		FileInputStream customerFile = new FileInputStream(
				new File(".\\source\\"+customerSheetName));

		XSSFWorkbook workbookCustomer = new XSSFWorkbook(customerFile);
		XSSFSheet customerSheet = workbookCustomer.getSheet("Customer");

		FileInputStream addressFile = new FileInputStream(
				new File(".\\source\\"+addressSheetName));

		XSSFWorkbook workbookAddress = new XSSFWorkbook(addressFile);
		XSSFSheet addressSheet = workbookAddress.getSheet("Address");

		int customerLastRow = customerSheet.getLastRowNum();

		Cell cell = null;
		String customerId = "";
		String uid = "";
		String tempUid = "";
		String customerEmail;
		
		for (int i = 4; i <= customerLastRow; i++) {

			if (customerSheet.getRow(i).getCell(3).getStringCellValue()!="") {
				
				customerEmail = customerSheet.getRow(i).getCell(3).toString();
				uid = customerSheet.getRow(i).getCell(1).toString();
				
				for (int j = 1; j <= row; j++) {
					
					customerId = "";

					if (("abc" + customerEmail)
							.equalsIgnoreCase(exportSheet.getRow(j).getCell(0).getStringCellValue())) {
						customerId = exportSheet.getRow(j).getCell(1).getStringCellValue();
						for (int k = 4; k <= addressSheet.getLastRowNum(); k++) {

							if (addressSheet.getRow(k).getCell(1).toString() != "") {

								tempUid = addressSheet.getRow(k).getCell(1).toString();

								if (uid.equals(tempUid)) {
									addressSheet.getRow(k).createCell(1);
									addressSheet.getRow(k).getCell(1).setCellValue(customerId);
									System.out.println("Inserting CustomerId In Address Sheet: "+customerId);
									
								}

							} 
						}
						
						cell = customerSheet.getRow(i).createCell(1);
	
						cell.setCellValue(customerId);
						System.out.println("Inserting CustomerId In Customer Sheet: "+customerId);
						break;
						
						}
						
						
						
					}

				}

		}

	//Deleting rows of Customer which does not have mapping in csv/Email
		for (int i = 4; i <= customerSheet.getLastRowNum(); i++) {
			try
			{
			if (customerSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC 
					|| customerSheet.getRow(i).getCell(1).toString() == "" 
					|| customerSheet.getRow(i).getCell(1).getCellType() == CellType.BLANK)
			{
				System.out.println("Removing records which does not have id mapping/Email");
				customerSheet.shiftRows(customerSheet.getRow(i).getRowNum() + 1, customerSheet.getLastRowNum() + 1, -1);
				i--;
			}
			}
			catch (NullPointerException e)
			{
				System.out.println("Null Pointer at row"+i);
				e.printStackTrace();
			}
		}

	//Deleting rows of Address which  does not have corresponding entry in Customer
	for(int i = 4;i<=addressSheet.getLastRowNum();i++)
	{
		if (addressSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC )
		{
			System.out.println("Removing records whic does not have corresponding value in customer");
			addressSheet.shiftRows(addressSheet.getRow(i).getRowNum() + 1, addressSheet.getLastRowNum() + 1, -1);
			i--;
		}
	}

	customerFile.close();
	try
	{
		FileOutputStream outFile = new FileOutputStream(new File(".\\UpdatedFiles\\newCustomer.xlsx"));
		workbookCustomer.write(outFile);
		outFile.close();
	}catch(
	Exception e)
	{
		System.out.println(e);

	}

	try
	{
		FileOutputStream outFileAddress = new FileOutputStream(new File(".\\UpdatedFiles\\newAddress.xlsx"));
		workbookAddress.write(outFileAddress);
		outFileAddress.close();
		System.out.println("Finished");
	}catch(
	Exception e)
	{
		System.out.println(e);

	}

	long end = System.currentTimeMillis();
	long elapsedTime = end - start;

	System.out.println("Total time taken: "+elapsedTime);
}}
