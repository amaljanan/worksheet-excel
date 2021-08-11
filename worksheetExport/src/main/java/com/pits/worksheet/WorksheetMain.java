package com.pits.worksheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.csv.QuoteMode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class WorksheetMain {
	
	static int customerValueRow = 4, addressValueRow = 4;

	public static void main(String[] args) throws Exception {

		//String sample = "26 av\\; mal\\; de lattre de tassigny";
		//System.out.println(sample.replaceAll("\\\\", ""));

		Cell cell = null;
		String customerId = "";
		String uid = "";
		String tempUid = "";
		String customerEmail;
		int customerCellCount = 10, addressCellCount = 14;
		int customerHeader = 2, addressHeader = 2;
		
		XSSFWorkbook workbookCustomer = null;
		XSSFSheet customerSheet;

		XSSFWorkbook workbookAddress = null;
		XSSFSheet addressSheet;

		List<CSVRecord> list = null;

		Scanner sc = new Scanner(System.in);


		try {
			list = importExportCSVFile(sc);

			workbookCustomer = importCustomer(sc);

			workbookAddress = importAddress(sc);

		} catch (Exception e) {
			e.printStackTrace();
		}

		customerSheet = workbookCustomer.getSheet("Customer");
		addressSheet = workbookAddress.getSheet("Address");
		
		
		XSSFWorkbook workbookDeletedCustomer = new XSSFWorkbook();
		XSSFSheet deletedCustomerSheet = workbookDeletedCustomer.createSheet("Deleted Customer");

		XSSFWorkbook workbookDeletedAddress = new XSSFWorkbook();
		XSSFSheet deletedAddressSheet = workbookDeletedAddress.createSheet("Deleted Address");

		long start = System.currentTimeMillis();
		
		
		//XSSFRow deleteRow;
		
		customerSheet = deleteGuestCustomer(customerSheet);
		
		for (int i = customerValueRow; i <= customerSheet.getLastRowNum(); i++) {

			if (customerSheet.getRow(i).getCell(3).getStringCellValue() != "") {

				customerEmail = customerSheet.getRow(i).getCell(3).toString();
				uid = customerSheet.getRow(i).getCell(1).toString();

				for (CSVRecord record : list) {

					customerId = "";

					if (("abc" + customerEmail).equalsIgnoreCase(record.get(0))) {
						customerId = record.get(1);
						for (int k = addressValueRow; k <= addressSheet.getLastRowNum(); k++) {

							if (addressSheet.getRow(k).getCell(1).toString() != "") {

								tempUid = addressSheet.getRow(k).getCell(1).toString();

								if (uid.equals(tempUid)) {
									addressSheet.getRow(k).createCell(1);
									addressSheet.getRow(k).getCell(1).setCellValue(customerId);
									System.out.println("Inserting CustomerId In Address Sheet: " + customerId);

								}
							}
						}

						cell = customerSheet.getRow(i).createCell(1);
						cell.setCellValue(customerId);

						System.out.println("Inserting CustomerId In Customer Sheet: " + customerId);

						// Changing customer type to guest
						customerSheet.getRow(i).createCell(7).setCellValue("Guest");
						break;

					}

				}

			}

		}
		customerSheet = deleteCustomerRow(customerSheet, deletedCustomerSheet);

		addressSheet = deleteAddressRow(addressSheet, deletedAddressSheet);

		try {
			FileOutputStream outFile = new FileOutputStream(new File(".\\UpdatedFiles\\newCustomer.xlsx"));
			workbookCustomer.write(outFile);
			outFile.close();

			FileOutputStream deleteCustomerOutFile = new FileOutputStream(
					new File(".\\UpdatedFiles\\deletedCustomer.xlsx"));
			workbookDeletedCustomer.write(deleteCustomerOutFile);
			deleteCustomerOutFile.close();

			workbookCustomer.close();
			workbookDeletedCustomer.close();

		} catch (Exception e) {
			System.out.println(e);

		}

		try {
			FileOutputStream outFileAddress = new FileOutputStream(new File(".\\UpdatedFiles\\newAddress.xlsx"));
			workbookAddress.write(outFileAddress);
			outFileAddress.close();

			FileOutputStream deleteAddressOutFile = new FileOutputStream(
					new File(".\\UpdatedFiles\\deletedAddress.xlsx"));
			workbookDeletedAddress.write(deleteAddressOutFile);
			deleteAddressOutFile.close();

			workbookAddress.close();
			workbookDeletedAddress.close();
			
			customerToImpex(customerSheet,customerCellCount,customerHeader);
			
			addressToImpex(addressSheet,addressCellCount,addressHeader);
			

			System.out.println("Finished");
		} catch (Exception e) {
			System.out.println(e);

		}

		long end = System.currentTimeMillis();
		long elapsedTime = end - start;

		System.out.println("Total time taken: " + elapsedTime);
		
		//System.out.println("Do You want to convert this to impex");
		//String inputFromUser = sc.nextLine();
		
		

	}
	
	
	//delete Guest from Customer
	private static XSSFSheet deleteGuestCustomer(XSSFSheet customerSheet) {
		for (int i = customerValueRow; i <= customerSheet.getLastRowNum(); i++) {
			try {
					if(customerSheet.getRow(i).getCell(7).toString().equalsIgnoreCase("Guest"))
					{
					customerSheet.shiftRows(customerSheet.getRow(i).getRowNum() + 1, customerSheet.getLastRowNum() + 1, -1);
					i--;
					}
				}
			catch (NullPointerException e) {
				System.out.println("Null Pointer at row" + i);
				e.printStackTrace();
				}
			} 
		return customerSheet;
	}

	private static XSSFSheet deleteAddressRow(XSSFSheet addressSheet, XSSFSheet deletedAddressSheet) {
		int insertRow = 1;
		XSSFRow deleteRow;
		
		deleteRow = deletedAddressSheet.createRow(0);
		deleteRow.createCell(0).setCellValue("uid");
		deleteRow.createCell(1).setCellValue("Reason for deletion");
		
		// Deleting rows of Address which does not have corresponding entry in Customer
		for (int i = addressValueRow; i <= addressSheet.getLastRowNum(); i++) {
			if (addressSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC) {
				System.out.println("Removing records which does not have corresponding value in customer");
				
				deleteRow = deletedAddressSheet.createRow(insertRow++);
				
				deleteRow.createCell(0).setCellValue(addressSheet.getRow(i).getCell(1).getNumericCellValue());
				
				if (addressSheet.getRow(i).getCell(1) == null || addressSheet.getRow(i).getCell(1).getCellType() == CellType.BLANK)
					deleteRow.createCell(1).setCellValue("Customer Id is not present ");
				else
					deleteRow.createCell(1).setCellValue("Corresponding Customer is not present in customer sheet ");
			
				addressSheet.shiftRows(addressSheet.getRow(i).getRowNum() + 1, addressSheet.getLastRowNum() + 1, -1);
				i--;
			}
		}

		return addressSheet;
	}

	private static XSSFSheet deleteCustomerRow(XSSFSheet customerSheet, XSSFSheet deletedCustomerSheet) {
		int insertRow = 1;
		XSSFRow deleteRow;
		
		deleteRow = deletedCustomerSheet.createRow(0);
		deleteRow.createCell(0).setCellValue("uid");
		deleteRow.createCell(1).setCellValue("email");
		deleteRow.createCell(2).setCellValue("Reason for deletion");
		
		// Deleting rows of Customer which does not have mapping in csv/Email
		for (int i = customerValueRow; i <= customerSheet.getLastRowNum(); i++) {
			try {
				if (customerSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC
						|| customerSheet.getRow(i).getCell(1).toString() == ""
						|| customerSheet.getRow(i).getCell(1).getCellType() == CellType.BLANK) {
					System.out.println("Removing records which does not have id mapping/Email");
					deleteRow = deletedCustomerSheet.createRow(insertRow++);
					
					deleteRow.createCell(0).setCellValue(customerSheet.getRow(i).getCell(1).toString());
					deleteRow.createCell(1).setCellValue(customerSheet.getRow(i).getCell(3).toString());
					
				      
				      if (customerSheet.getRow(i).getCell(3) == null || customerSheet.getRow(i).getCell(3).getCellType() == CellType.BLANK)
							deleteRow.createCell(2).setCellValue("No Email Id Present for this record");
						else
							deleteRow.createCell(2).setCellValue("No Cutomer Id mapping in Export Sheet");
				
					customerSheet.shiftRows(customerSheet.getRow(i).getRowNum() + 1, customerSheet.getLastRowNum() + 1, -1);
					i--;
				}
			} catch (NullPointerException e) {
				System.out.println("Null Pointer at row" + i);
				e.printStackTrace();
			}
		}

		return customerSheet;
	}

	private static XSSFWorkbook importCustomer(Scanner sc) throws IOException {
		System.out.println("Enter Customer Sheet name with extension:");
		String customerSheetName = sc.nextLine();

		FileInputStream customerFile = new FileInputStream(new File(".\\source\\" + customerSheetName));

		XSSFWorkbook workbookCustomer = new XSSFWorkbook(customerFile);
		return workbookCustomer;

	}

	private static XSSFWorkbook importAddress(Scanner sc) throws IOException {
		System.out.println("Enter Address Sheet name with Extension:");
		String addressSheetName = sc.nextLine();

		FileInputStream addressFile = new FileInputStream(new File(".\\source\\" + addressSheetName));
		XSSFWorkbook workbookAddress = new XSSFWorkbook(addressFile);

		return workbookAddress;
	}

	private static List<CSVRecord> importExportCSVFile(Scanner sc) throws IOException {

		String exportSheetName = null;

		System.out.println("Enter Export Sheet Name with extension:");
		exportSheetName = sc.nextLine();

		CSVParser exportCSVParser = new CSVParser(new FileReader(new File(".\\source\\" + exportSheetName)),
				CSVFormat.DEFAULT);
		// exportCSVParser.close();
		return exportCSVParser.getRecords();
	}
	
	
	
	
	private static void customerToImpex(XSSFSheet customerSheet, int customerCellCount, int customerHeader) throws Exception
	{


		FileOutputStream sampleCsv = new FileOutputStream(new File(".\\UpdatedFiles\\CustomerImpex.impex"));
		CSVPrinter csvPrinter = null; 
		csvPrinter = new CSVPrinter(new OutputStreamWriter(sampleCsv), CSVFormat.DEFAULT.withDelimiter(';').withTrim());
		
		for(int i=0;i<=customerCellCount;i++)
		{
			if(customerSheet.getRow(customerHeader).getCell(i)!=null)
			csvPrinter.print(customerSheet.getRow(customerHeader).getCell(i).toString());
		}
			csvPrinter.print(null);
		
        for(int i=customerValueRow;i<=customerSheet.getLastRowNum();i++) {                
        	
        		csvPrinter.println();
        		//csvPrinter.print(null);
        	for (int j =0;j<=customerCellCount;j++) {
            	
        		if(customerSheet.getRow(i).getCell(j)!=null)
        		{
        			if(customerSheet.getRow(i).getCell(j).getCellType()==CellType.NUMERIC)
        				csvPrinter.print(customerSheet.getRow(i).getCell(j).toString().split("\\.")[0]);
        			else
        				csvPrinter.print(customerSheet.getRow(i).getCell(j).toString());
        		}
        		else
        			csvPrinter.print(null);
            	
        	}
        		csvPrinter.print(null);
           
        }   
		
		csvPrinter.close();

	}
	
	
	private static void addressToImpex(XSSFSheet addressSheet, int addressCellCount, int addressHeader) throws Exception
	{


		FileOutputStream sampleCsv = new FileOutputStream(new File(".\\UpdatedFiles\\AddressImpex.impex"));
		CSVPrinter csvPrinter = null; 
		csvPrinter = new CSVPrinter(new OutputStreamWriter(sampleCsv), CSVFormat.DEFAULT.withDelimiter(';').withTrim().withQuoteMode(QuoteMode.MINIMAL));
		
		for(int i=0;i<=addressCellCount;i++)
		{
			if(addressSheet.getRow(addressHeader).getCell(i)!=null)
			csvPrinter.print(addressSheet.getRow(addressHeader).getCell(i).toString());
		}
			csvPrinter.print(null);
		
        for(int i=addressValueRow;i<=addressSheet.getLastRowNum();i++) {                
        	
        		csvPrinter.println();
        		//csvPrinter.print(null);
        	for (int j =0;j<=addressCellCount;j++) {
            	
        		if(null != addressSheet.getRow(i).getCell(j) && addressSheet.getRow(i).getCell(j).toString()!="")
        		{
        			/*if(j==5 || j==6 || j==8)
        			{
        				String value = addressSheet.getRow(i).getCell(j).toString().replaceAll("\"","'");
        				value = value.replaceAll("\\\\", "");
        				value = "\"" + value + "\"";
        				csvPrinter.print(value);
        			}*/
        			if(addressSheet.getRow(i).getCell(j).getCellType()==CellType.NUMERIC)
        				csvPrinter.print(addressSheet.getRow(i).getCell(j).toString().split("\\.")[0]);
        			else
        				csvPrinter.print(addressSheet.getRow(i).getCell(j).toString());
        		}
        		else
        		{
        			csvPrinter.print(null);
        		}
            	
        	} 
        		csvPrinter.print(null);
           
        }   
		
		csvPrinter.close();

	}

}
