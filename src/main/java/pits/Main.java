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
import java.util.Scanner;

public class Main {

  public static void main(String[] args) {

    try {

      Scanner scanner = new Scanner(System.in);

      XSSFWorkbook customerWorkbook = importCustomerWorkBook(scanner);

      XSSFWorkbook addressWorkbook = importAddressWorkBook(scanner);

      List<CSVRecord> list = importExportCSVFile(scanner);

      long start = System.currentTimeMillis();

      mappingCustomerWorkbook(customerWorkbook, list);

      mappingAddressWorkbook(customerWorkbook, addressWorkbook);

      cleaningUpWorkBooks(customerWorkbook, addressWorkbook);

      exportingFinalCustomerWorkbook(customerWorkbook);

      exportingFinalAddressWorkbook(addressWorkbook);

      long end = System.currentTimeMillis();

      System.out.println("Excel Work Book TookT = " + (end - start) + "ms");

      System.out.println("Do you wish to create impex files ? (y/n)");
      if (scanner.nextLine().equalsIgnoreCase("y")) {

        createCustomerImpexFile(customerWorkbook);

        createAddressImpexFile(addressWorkbook);
      }

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static XSSFWorkbook importCustomerWorkBook(Scanner scanner) throws IOException {
    System.out.println("Enter Customer Sheet name with extension : ");
    String customerSheetName = scanner.nextLine();

    FileInputStream customerFileInputStream = new FileInputStream("./Source/" + customerSheetName);
    // customerFileInputStream.close();

    return new XSSFWorkbook(customerFileInputStream);
  }

  private static XSSFWorkbook importAddressWorkBook(Scanner scanner) throws IOException {
    System.out.println("Enter Address Sheet name with Extension : ");
    String addressSheetName = scanner.nextLine();

    FileInputStream addressFileInputStream = new FileInputStream("./Source/" + addressSheetName);
    // addressFileInputStream.close();
    return new XSSFWorkbook(addressFileInputStream);
  }

  private static List<CSVRecord> importExportCSVFile(Scanner scanner) throws IOException {

    System.out.println("Enter Export Sheet Name with extension : ");
    String exportSheetName = scanner.nextLine();

    CSVParser exportCSVParser =
        new CSVParser(new FileReader("./Source/" + exportSheetName), CSVFormat.DEFAULT);
    // exportCSVParser.close();
    return exportCSVParser.getRecords();
  }

  private static void mappingCustomerWorkbook(XSSFWorkbook customerWorkbook, List<CSVRecord> list) {

    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    for (int i = 0; i <= customerSheet.getLastRowNum(); i++) {

      if (null != customerSheet.getRow(i).getCell(3)
          && !customerSheet.getRow(i).getCell(3).getStringCellValue().isEmpty()
          && !customerSheet.getRow(i).getCell(3).getStringCellValue().equals("")) {
        for (CSVRecord record : list) {
          if (customerSheet
              .getRow(i)
              .getCell(3)
              .getStringCellValue()
              .equalsIgnoreCase(record.get(0).substring(3))) {

            String uid = customerSheet.getRow(i).getCell(1).toString().split("\\.")[0];
            customerSheet
                .getRow(i)
                .createCell(1)
                .setCellValue(uid.concat("##").concat(record.get(1)));
            break;
          }
        }
      }
    }
  }

  private static void mappingAddressWorkbook(
      XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook) {

    XSSFSheet addressSheet = addressWorkbook.getSheet("Address");
    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    for (int i = 0; i <= addressSheet.getLastRowNum(); i++) {
      if (null != addressSheet.getRow(i).getCell(1)
          && !addressSheet.getRow(i).getCell(1).toString().equals("")) {

        String addressSheetUid = addressSheet.getRow(i).getCell(1).toString().split("\\.")[0];

        for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {
          if (null != customerSheet.getRow(j).getCell(1)
              && customerSheet.getRow(j).getCell(1).toString().contains("##")) {

            String[] customerSheetId = customerSheet.getRow(j).getCell(1).toString().split("##");
            String customerUid = customerSheetId[0];

            if (customerUid.equalsIgnoreCase(addressSheetUid)) {
              String id = customerSheetId[1];
              addressSheet.getRow(i).createCell(1).setCellValue(id);
              System.out.println("Mapped for Address WorkBook with Customer id =" + id);
              break;
            }
          }
        }
      }
    }

    for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {

      if (null != customerSheet.getRow(j).getCell(1)
          && customerSheet.getRow(j).getCell(1).toString().contains("##")) {

        String[] customerSheetId = customerSheet.getRow(j).getCell(1).toString().split("##");
        String id = customerSheetId[1];
        customerSheet.getRow(j).createCell(1).setCellValue(id);
        System.out.println("Mapped for Customer WorkBook with Customer id = " + id);
      }

      if (null != customerSheet.getRow(j).getCell(7)
          && customerSheet.getRow(j).getCell(7).toString().equalsIgnoreCase("Registred")) {
        customerSheet.getRow(j).createCell(7).setCellValue("Guest");
      }
    }
  }

  private static void cleaningUpWorkBooks(
      XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook) throws IOException {

    XSSFWorkbook deletedEntriesWorkBook = new XSSFWorkbook();

    removeInvalidEntriesFromCustomerWorkbook(customerWorkbook, deletedEntriesWorkBook);

    removeInvalidEntriesFromAddress(addressWorkbook, deletedEntriesWorkBook);

    FileOutputStream fos = new FileOutputStream("./Final/DeletedRecords.xlsx");
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

      if (null != customerSheet.getRow(j).getCell(1)
          && (customerSheet.getRow(j).getCell(1).getCellType() == CellType.NUMERIC
              || customerSheet.getRow(j).getCell(1).toString().equals("")
              || customerSheet.getRow(j).getCell(1).getCellType() == CellType.BLANK)) {

        XSSFRow deletedRow = deletedCustomerSheet.createRow(deleteSheetRowNumber++);

        deletedRow.createCell(0).setCellValue(customerSheet.getRow(j).getCell(1).getNumericCellValue());

        if(null != customerSheet.getRow(j).getCell(3))
        deletedRow.createCell(1).setCellValue(customerSheet.getRow(j).getCell(3).toString());

        /*CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyCellStyle(false);

        deletedRow.copyRowFrom(customerSheet.getRow(j), policy);*/

        System.out.println(
            "Removed Invalid Entry with Customer id =" + customerSheet.getRow(j).getCell(1));
        if (customerSheet.getRow(j).getCell(3) == null
            || customerSheet.getRow(j).getCell(3).getCellType() == CellType.BLANK)
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

      if (null == addressSheet.getRow(j).getCell(1)
          || addressSheet.getRow(j).getCell(1).getCellType() == CellType.NUMERIC) {

        XSSFRow deletedRow = deletedAddressSheet.createRow(deleteSheetRowNumber++);

       // deletedRow.createCell(1).setCellValue(addressSheet.getRow(j).getCell(1).toString());
        deletedRow.createCell(0).setCellValue(addressSheet.getRow(j).getCell(1).getNumericCellValue());

        /*CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyCellStyle(false);
        try {
          deletedRow.copyRowFrom(addressSheet.getRow(j), policy);
        } catch (Exception e) {
          System.out.println(
                  "Exception while logging deleted address entry for customerId = "
                          + addressSheet.getRow(j).getCell(1));
        }*/

        System.out.println(
            "Removed Invalid Address with Customer id =" + addressSheet.getRow(j).getCell(1));

        if (null == addressSheet.getRow(j).getCell(1))
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

    FileOutputStream addressFileOutputStream = new FileOutputStream("./Final/Address.xlsx");
    addressWorkbook.write(addressFileOutputStream);

    addressFileOutputStream.close();
  }

  private static void exportingFinalCustomerWorkbook(XSSFWorkbook customerWorkbook)
      throws IOException {
    FileOutputStream customerFileOutputStream = new FileOutputStream("./Final/Customer.xlsx");
    customerWorkbook.write(customerFileOutputStream);
    customerFileOutputStream.close();
  }

  private static void createCustomerImpexFile(XSSFWorkbook customerWorkbook) {

    CSVPrinter csvPrinter = null;
    try {

      csvPrinter =
          new CSVPrinter(
              new FileWriter("./Final/CustomerImpex.impex"),
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
          for (int j = 0; j <= 9; j++) {
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
              new FileWriter("./Final/AddressImpex.impex"),
              CSVFormat.EXCEL
                  .withDelimiter(';')
                  .withTrim()
                  .withEscape('\\')
                  .withQuoteMode(QuoteMode.NONE));

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
          for (int j = 0; j <= 13; j++) {
            if (null != row.getCell(j) && !row.getCell(j).toString().equalsIgnoreCase("")) {
              String value = row.getCell(j).toString();
              if (row.getCell(j).getCellType() == CellType.NUMERIC && value.contains(".")) {
                value = value.split("\\.")[0];
              }
              if (j == 5 || j == 6 || j == 8) {
                value = value.replaceAll("\"", "'");
                value = "\"" + value + "\"";
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
