package com.pits.worksheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorksheetMain {

	public static void main(String[] args) {

		XSSFWorkbook workbookCustomer = null;
		XSSFSheet customerSheet;

		XSSFWorkbook workbookAddress = null;
		XSSFSheet addressSheet;

		List<CSVRecord> list = null;

		Scanner sc = new Scanner(System.in);

		long start = System.currentTimeMillis();

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

		Cell cell = null;
		String customerId = "";
		String uid = "";
		String tempUid = "";
		String customerEmail;

		for (int i = 4; i <= customerSheet.getLastRowNum(); i++) {

			if (customerSheet.getRow(i).getCell(3).getStringCellValue() != "") {

				customerEmail = customerSheet.getRow(i).getCell(3).toString();
				uid = customerSheet.getRow(i).getCell(1).toString();

				for (CSVRecord record : list) {

					customerId = "";

					if (("abc" + customerEmail).equalsIgnoreCase(record.get(0))) {
						customerId = record.get(1);
						for (int k = 4; k <= addressSheet.getLastRowNum(); k++) {

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

			System.out.println("Finished");
		} catch (Exception e) {
			System.out.println(e);

		}

		long end = System.currentTimeMillis();
		long elapsedTime = end - start;

		System.out.println("Total time taken: " + elapsedTime);

	}

	private static XSSFSheet deleteAddressRow(XSSFSheet addressSheet, XSSFSheet deletedAddressSheet) {
		int insertRow = 0;
		Row deleteRow;
		// Deleting rows of Address which does not have corresponding entry in Customer
		for (int i = 4; i <= addressSheet.getLastRowNum(); i++) {
			if (addressSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC) {
				System.out.println("Removing records which does not have corresponding value in customer");
				deleteRow = deletedAddressSheet.createRow(insertRow++);
				for (int j = 0; j <= addressSheet.getRow(4).getLastCellNum(); j++) {
					Cell mycell;
					mycell = deleteRow.createCell(j, CellType.STRING);
					if (addressSheet.getRow(i).getCell(j) != null) {
						// System.out.println(customerSheet.getRow(i).getCell(j).toString());
						mycell.setCellValue(addressSheet.getRow(i).getCell(j).toString());
					}
					if (addressSheet.getRow(i).getCell(1) == null)
						deleteRow.createCell(15).setCellValue("Reason for deletion: Customer Id is not present ");
					else
						deleteRow.createCell(15).setCellValue(
								"Reason for deletion: Corresponding Customer is not present in customer sheet ");
				}
				addressSheet.shiftRows(addressSheet.getRow(i).getRowNum() + 1, addressSheet.getLastRowNum() + 1, -1);
				i--;
			}
		}

		return addressSheet;
	}

	private static XSSFSheet deleteCustomerRow(XSSFSheet customerSheet, XSSFSheet deletedCustomerSheet) {
		int insertRow = 0;
		Row deleteRow;
		// Deleting rows of Customer which does not have mapping in csv/Email
		for (int i = 4; i <= customerSheet.getLastRowNum(); i++) {
			try {
				if (customerSheet.getRow(i).getCell(1).getCellType() == CellType.NUMERIC
						|| customerSheet.getRow(i).getCell(1).toString() == ""
						|| customerSheet.getRow(i).getCell(1).getCellType() == CellType.BLANK) {
					System.out.println("Removing records which does not have id mapping/Email");
					deleteRow = deletedCustomerSheet.createRow(insertRow++);
					for (int j = 0; j <= customerSheet.getRow(4).getLastCellNum(); j++) {
						Cell mycell;
						mycell = deleteRow.createCell(j, CellType.STRING);
						if (customerSheet.getRow(i).getCell(j) != null) {
							// System.out.println(customerSheet.getRow(i).getCell(j).toString());
							mycell.setCellValue(customerSheet.getRow(i).getCell(j).toString());
						}
						if (customerSheet.getRow(i).getCell(3) == null
								|| customerSheet.getRow(i).getCell(3).getCellType() == CellType.BLANK)
							deleteRow.createCell(11)
									.setCellValue("Reason for deletion : No Email Id Present for this record");
						else
							deleteRow.createCell(11)
									.setCellValue("Reason for deletion : No Cutomer Id mapping in Export Sheet");
					}

					customerSheet.shiftRows(customerSheet.getRow(i).getRowNum() + 1, customerSheet.getLastRowNum() + 1,
							-1);
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

}