package pits;

import org.apache.commons.csv.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Scanner;

public class Main {

  private static final int customerSheetColumnCount = 9;
  private static final int customerSheetUidIndex = 2;
  private static final int customerSheetEmailIndex = 4;
  private static final int customerSheetTypeCodeIndex = 8;


  private static final int addressSheetColumnCount = 14;
  private static final int addressSheetCustomerUidIndex = 2;



  public static void main(String[] args) {

    try {

      Scanner scanner = new Scanner(System.in);
      boolean isEnviornmentUAT = true;

      XSSFWorkbook customerWorkbook = importCustomerWorkBook(scanner);

      XSSFWorkbook addressWorkbook = importAddressWorkBook(scanner);

      List<CSVRecord> list = importExportCSVFile(scanner);

      System.out.println("Select the environment : (1/2)");
      System.out.println("1. UAT ");
      System.out.println("2. PROD");

      if (scanner.nextLine().equalsIgnoreCase("2")) {
        isEnviornmentUAT = false;
      }

      long start = System.currentTimeMillis();

      deleteGuestCustomer(customerWorkbook);

      mappingCustomerWorkbook(customerWorkbook, list);

      mappingAddressWorkbook(customerWorkbook, addressWorkbook, isEnviornmentUAT);

      cleaningUpWorkBooks(customerWorkbook, addressWorkbook);

      exportingFinalCustomerWorkbook(customerWorkbook);

      exportingFinalAddressWorkbook(addressWorkbook);

      long end = System.currentTimeMillis();

      System.out.println("Excel Work Book TookT = " + (end - start) + "ms");

      System.out.println("Do you wish to create impex files ? (y/n)");
      if (scanner.nextLine().equalsIgnoreCase("y")) {
        System.out.println("Creating Customer Impex file..");
        createCustomerImpexFile(customerWorkbook);
        System.out.println("Creating Address Impex file..");
        createAddressImpexFile(addressWorkbook);
        System.out.println("Impex Files created");
      }

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static XSSFWorkbook importCustomerWorkBook(Scanner scanner) throws IOException {
    System.out.println("Enter Customer Sheet name with extension : ");
    String customerSheetName = scanner.nextLine();

    FileInputStream customerFileInputStream =
        new FileInputStream("./Source Folder/" + customerSheetName);
    // customerFileInputStream.close();

    return new XSSFWorkbook(customerFileInputStream);
  }

  private static XSSFWorkbook importAddressWorkBook(Scanner scanner) throws IOException {
    System.out.println("Enter Address Sheet name with Extension : ");
    String addressSheetName = scanner.nextLine();

    FileInputStream addressFileInputStream =
        new FileInputStream("./Source Folder/" + addressSheetName);
    // addressFileInputStream.close();
    return new XSSFWorkbook(addressFileInputStream);
  }

  private static List<CSVRecord> importExportCSVFile(Scanner scanner) throws IOException {

    System.out.println("Enter Export Sheet Name with extension : ");
    String exportSheetName = scanner.nextLine();

    CSVParser exportCSVParser =
        new CSVParser(new FileReader("./Source Folder/" + exportSheetName), CSVFormat.DEFAULT);
    // exportCSVParser.close();
    return exportCSVParser.getRecords();
  }

  private static void deleteGuestCustomer(XSSFWorkbook customerWorkbook) {

    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    System.out.println("Removing Guest Customer from Customer Work Book...");
    for (int i = 0; i <= customerSheet.getLastRowNum(); i++) {
      try {
        if (null != customerSheet.getRow(i).getCell(customerSheetTypeCodeIndex)
                && !customerSheet.getRow(i).getCell(customerSheetTypeCodeIndex).getStringCellValue().isEmpty()
                && !customerSheet.getRow(i).getCell(customerSheetTypeCodeIndex).getStringCellValue().equals("")
                && customerSheet.getRow(i).getCell(customerSheetTypeCodeIndex).toString().equalsIgnoreCase("Guest")) {
          customerSheet.shiftRows(
                  customerSheet.getRow(i).getRowNum() + 1, customerSheet.getLastRowNum() + 1, -1);
          i--;
        }
      } catch (NullPointerException e) {
        System.out.println("Null Pointer while delete ing Guest Customer at row" + i);
        e.printStackTrace();
      }
    }
  }

  private static void mappingCustomerWorkbook(XSSFWorkbook customerWorkbook, List<CSVRecord> list) {

    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    for (int i = 0; i <= customerSheet.getLastRowNum(); i++) {

      if (null != customerSheet.getRow(i).getCell(customerSheetEmailIndex)
          && !customerSheet.getRow(i).getCell(customerSheetEmailIndex).getStringCellValue().isEmpty()
          && !customerSheet.getRow(i).getCell(customerSheetEmailIndex).getStringCellValue().equals("")) {
        for (CSVRecord record : list) {
          if (customerSheet
              .getRow(i)
              .getCell(customerSheetEmailIndex)
              .getStringCellValue()
              .equalsIgnoreCase(record.get(0).substring(3))) {

            String uid = customerSheet.getRow(i).getCell(customerSheetUidIndex).toString().split("\\.")[0];
            customerSheet
                .getRow(i)
                .createCell(customerSheetUidIndex)
                .setCellValue(uid.concat("##").concat(record.get(1)));
            break;
          }
        }
      }
    }
  }

  private static void mappingAddressWorkbook(
      XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook, Boolean isEnviornmentUAT) {

    XSSFSheet addressSheet = addressWorkbook.getSheet("Address");
    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    for (int i = 0; i <= addressSheet.getLastRowNum(); i++) {
      if (null != addressSheet.getRow(i).getCell(addressSheetCustomerUidIndex)
          && !addressSheet.getRow(i).getCell(addressSheetCustomerUidIndex).toString().equals("")) {

        String addressSheetUid = addressSheet.getRow(i).getCell(addressSheetCustomerUidIndex).toString().split("\\.")[0];

        for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {
          if (null != customerSheet.getRow(j).getCell(customerSheetUidIndex)
              && customerSheet.getRow(j).getCell(customerSheetUidIndex).toString().contains("##")) {

            String[] customerSheetId = customerSheet.getRow(j).getCell(customerSheetUidIndex).toString().split("##");
            String customerUid = customerSheetId[0];

            if (customerUid.equalsIgnoreCase(addressSheetUid)) {
              String id = customerSheetId[1];
              addressSheet.getRow(i).createCell(addressSheetCustomerUidIndex).setCellValue(id);
              System.out.println("Mapped for Address WorkBook with Customer id =" + id);
              break;
            }
          }
        }
      }
    }

    for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {

      if (null != customerSheet.getRow(j).getCell(customerSheetUidIndex)
          && customerSheet.getRow(j).getCell(customerSheetUidIndex).toString().contains("##")) {

        String[] customerSheetId = customerSheet.getRow(j).getCell(customerSheetUidIndex).toString().split("##");
        String id = customerSheetId[1];
        customerSheet.getRow(j).createCell(customerSheetUidIndex).setCellValue(id);
        System.out.println("Mapped for Customer WorkBook with Customer id = " + id);
      }

      if (null != customerSheet.getRow(j).getCell(customerSheetTypeCodeIndex)
          && customerSheet.getRow(j).getCell(customerSheetTypeCodeIndex).toString().equalsIgnoreCase("Registred")) {
        customerSheet.getRow(j).createCell(customerSheetTypeCodeIndex).setCellValue("Guest");
      }
      if (j != 2
          && j != 3
          && null != customerSheet.getRow(j).getCell(customerSheetEmailIndex)
          && customerSheet.getRow(j).getCell(customerSheetEmailIndex).getCellType() != CellType.BLANK) {
        String email = customerSheet.getRow(j).getCell(customerSheetEmailIndex).toString();
        if (isEnviornmentUAT) email = "abc".concat(email.toLowerCase(Locale.ROOT));
        else email = email.toLowerCase(Locale.ROOT);
        customerSheet.getRow(j).createCell(customerSheetEmailIndex).setCellValue(email);
      }
    }
  }

  private static void cleaningUpWorkBooks(
      XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook) throws IOException {

    XSSFWorkbook deletedEntriesWorkBook = new XSSFWorkbook();

    removeInvalidEntriesFromCustomerWorkbook(customerWorkbook, deletedEntriesWorkBook);

    removeInvalidEntriesFromAddress(addressWorkbook, deletedEntriesWorkBook);

    FileOutputStream fos = new FileOutputStream("./Target Folder/DeletedRecords.xlsx");
    deletedEntriesWorkBook.write(fos);
    fos.close();
  }

  private static void removeInvalidEntriesFromCustomerWorkbook(
      XSSFWorkbook customerWorkbook, XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet deletedCustomerSheet = deletedEntriesWorkBook.createSheet("Deleted Customer");

    int deleteSheetRowNumber = 0;

    XSSFRow headerRow2 = deletedCustomerSheet.createRow(deleteSheetRowNumber);

    headerRow2.createCell(0).setCellValue("Uid");
    headerRow2.createCell(1).setCellValue("Email");
    headerRow2.createCell(3).setCellValue("Reason");

    deleteSheetRowNumber++;

    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");
    for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {

      if (null != customerSheet.getRow(j).getCell(customerSheetUidIndex)
          && (customerSheet.getRow(j).getCell(customerSheetUidIndex).getCellType() == CellType.NUMERIC
              || customerSheet.getRow(j).getCell(customerSheetUidIndex).toString().equals("")
              || customerSheet.getRow(j).getCell(customerSheetUidIndex).getCellType() == CellType.BLANK)) {

        XSSFRow deletedRow = deletedCustomerSheet.createRow(deleteSheetRowNumber++);

        deletedRow
            .createCell(0)
            .setCellValue(customerSheet.getRow(j).getCell(customerSheetUidIndex).getNumericCellValue());

        if (null != customerSheet.getRow(j).getCell(customerSheetEmailIndex))
          deletedRow.createCell(1).setCellValue(customerSheet.getRow(j).getCell(customerSheetEmailIndex).toString());

        System.out.println(
            "Removed Invalid Entry with Customer id =" + customerSheet.getRow(j).getCell(1));
        if (customerSheet.getRow(j).getCell(customerSheetEmailIndex) == null
            || customerSheet.getRow(j).getCell(customerSheetEmailIndex).getCellType() == CellType.BLANK)
          deletedRow
              .createCell(3)
              .setCellValue("Reason for deletion : No Email Id Present for this record");
        else
          deletedRow
              .createCell(3)
              .setCellValue("Reason for deletion : No Customer Id mapping in Export Sheet");

        customerSheet.shiftRows(
            customerSheet.getRow(j).getRowNum() + 1, customerSheet.getLastRowNum() + 1, -1);
        j--;
      }
    }
  }

  private static void removeInvalidEntriesFromAddress(
      XSSFWorkbook addressWorkbook, XSSFWorkbook deletedEntriesWorkBook) {
    XSSFSheet addressSheet = addressWorkbook.getSheet("Address");

    XSSFSheet deletedAddressSheet = deletedEntriesWorkBook.createSheet("Deleted Address");

    int deleteSheetRowNumber = 0;

    XSSFRow headerRow2 = deletedAddressSheet.createRow(deleteSheetRowNumber);

    headerRow2.createCell(0).setCellValue("Customer uid ");
    headerRow2.createCell(2).setCellValue("Reason");

    deleteSheetRowNumber++;

    for (int j = 0; j <= addressSheet.getLastRowNum(); j++) {

      if (null == addressSheet.getRow(j).getCell(addressSheetCustomerUidIndex)
          || addressSheet.getRow(j).getCell(addressSheetCustomerUidIndex).getCellType() == CellType.NUMERIC) {

        XSSFRow deletedRow = deletedAddressSheet.createRow(deleteSheetRowNumber++);

        deletedRow
            .createCell(0)
            .setCellValue(addressSheet.getRow(j).getCell(addressSheetCustomerUidIndex).getNumericCellValue());

        System.out.println(
            "Removed Invalid Address with Customer id =" + addressSheet.getRow(j).getCell(addressSheetCustomerUidIndex));

        if (null == addressSheet.getRow(j).getCell(addressSheetCustomerUidIndex))
          deletedRow.createCell(2).setCellValue("Reason for deletion: Customer Id is not present ");
        else
          deletedRow
              .createCell(2)
              .setCellValue("Reason for deletion : No Customer Id mapping in Export Sheet");

        addressSheet.shiftRows(
            addressSheet.getRow(j).getRowNum() + 1, addressSheet.getLastRowNum() + 1, -1);
        j--;
      }
    }
  }

  private static void exportingFinalAddressWorkbook(XSSFWorkbook addressWorkbook)
      throws IOException {

    FileOutputStream addressFileOutputStream = new FileOutputStream("./Target Folder/Address.xlsx");
    addressWorkbook.write(addressFileOutputStream);

    addressFileOutputStream.close();
  }

  private static void exportingFinalCustomerWorkbook(XSSFWorkbook customerWorkbook)
      throws IOException {
    FileOutputStream customerFileOutputStream =
        new FileOutputStream("./Target Folder/Customer.xlsx");
    customerWorkbook.write(customerFileOutputStream);
    customerFileOutputStream.close();
  }

  private static void createCustomerImpexFile(XSSFWorkbook customerWorkbook) {

    CSVPrinter csvPrinter = null;
    try {

      csvPrinter =
          new CSVPrinter(
              new FileWriter("./Target Folder/CustomerImpex.impex"),
              CSVFormat.EXCEL.withDelimiter(';').withTrim());

      if (customerWorkbook != null) {
        XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

        Row headerRow = customerSheet.getRow(2);
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          if (null != cell && !cell.toString().equalsIgnoreCase("")) {
            csvPrinter.print(cell.toString());
          }
        }
        csvPrinter.print(null);
        csvPrinter.println();
        for (int i = 4; i <= customerSheet.getLastRowNum(); i++) {
          Row row = customerSheet.getRow(i);
          for (int j = 0; j <= customerSheetColumnCount; j++) {
            if (null != row.getCell(j)) {
              String value = row.getCell(j).toString();
              if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                value = value.split("\\.")[0];
              }
              csvPrinter.print(value);
            } else csvPrinter.print(null);
          }
          csvPrinter.print(null);
          csvPrinter.println();
        }
      }

    } catch (Exception e) {
      System.out.println("Failed to write Customer CSV file to output stream : ");
      e.printStackTrace();
    } finally {
      try {
        if (csvPrinter != null) {
          csvPrinter.flush(); // Flush and close CSVPrinter
          csvPrinter.close();
        }
      } catch (IOException ioe) {
        System.out.println("Error when closing CSV Printer");
      }
    }
  }

  private static void createAddressImpexFile(XSSFWorkbook addressWorkbook) {

    CSVPrinter csvPrinter = null;
    try {

      csvPrinter =
          new CSVPrinter(
              new FileWriter("./Target Folder/AddressImpex.impex"),
              CSVFormat.EXCEL.withDelimiter(';').withTrim().withQuoteMode(QuoteMode.MINIMAL));

      if (addressWorkbook != null) {
        XSSFSheet addressSheet = addressWorkbook.getSheet("Address");

        Row headerRow = addressSheet.getRow(2);
        Iterator<Cell> cellIterator = headerRow.cellIterator();
        while (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          if (null != cell && !cell.toString().equalsIgnoreCase("")) {
            csvPrinter.print(cell.toString());
          }
        }
        csvPrinter.print(null);
        csvPrinter.println();
        for (int i = 4; i <= addressSheet.getLastRowNum(); i++) {
          Row row = addressSheet.getRow(i);
          for (int j = 0; j <= addressSheetColumnCount; j++) {
            if (null != row.getCell(j) && !row.getCell(j).toString().equalsIgnoreCase("")) {
              String value = row.getCell(j).toString();
              if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                value = value.split("\\.")[0];
              }
              csvPrinter.print(value);
            } else csvPrinter.print(null);
          }
          csvPrinter.print(null);
          csvPrinter.println();
        }
      }

    } catch (Exception e) {
      System.out.println("Failed to write AddressCSV file to output stream : ");
      e.printStackTrace();
    } finally {
      try {
        if (csvPrinter != null) {
          csvPrinter.flush(); // Flush and close CSVPrinter
          csvPrinter.close();
        }
      } catch (IOException ioe) {
        System.out.println("Error when closing CSV Printer");
      }
    }
  }
}
