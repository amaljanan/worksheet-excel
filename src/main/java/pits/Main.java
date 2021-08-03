package pits;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
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

      System.out.println("Time took = " + (end - start));

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void cleaningUpWorkBooks(
          XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook) throws IOException {

    XSSFWorkbook deletedEntriesWorkBook = new XSSFWorkbook();

    removeInvalidEntriesFromCustomerWorkbook(customerWorkbook, deletedEntriesWorkBook);

    removeInvalidEntriesFromAddress(addressWorkbook, deletedEntriesWorkBook);

    FileOutputStream fos = new FileOutputStream(new File("./Final/Customerlog.xlsx"));
    deletedEntriesWorkBook.write(fos);
    fos.close();
  }

  private static void removeInvalidEntriesFromCustomerWorkbook(
          XSSFWorkbook customerWorkbook, XSSFWorkbook deletedEntriesWorkBook) {

    XSSFSheet deletedCustomerSheet = deletedEntriesWorkBook.createSheet("Deleted Customer");

    int deleteSheetRowNumber = 0;

    XSSFRow headerRow = deletedCustomerSheet.createRow(deleteSheetRowNumber);

    headerRow.createCell(1).setCellValue("MEMBER_ID");
    headerRow.createCell(3).setCellValue("EMAIL");
    headerRow.createCell(4).setCellValue("LASTNAME FIRSTNAME");
    headerRow.createCell(6).setCellValue("PHONE1");
    headerRow.createCell(7).setCellValue("typeclient");
    headerRow.createCell(8).setCellValue("EMAIL1");
    headerRow.createCell(9).setCellValue("MEMBER_ID");
    headerRow.createCell(10).setCellValue("ADDRESS_ID");

    deleteSheetRowNumber++;

    XSSFRow headerRow2 = deletedCustomerSheet.createRow(deleteSheetRowNumber);

    headerRow2.createCell(1).setCellValue("Uid");
    headerRow2.createCell(2).setCellValue("Customer Id(not to be filled)");
    headerRow2.createCell(3).setCellValue("Email");
    headerRow2.createCell(4).setCellValue("Name");
    headerRow2.createCell(5).setCellValue("Mobile Number");
    headerRow2.createCell(6).setCellValue("Customer Type");
    headerRow2.createCell(8).setCellValue("original customer ref number");
    headerRow2.createCell(9).setCellValue("address ID");
    headerRow2.createCell(11).setCellValue("REASON");

    deleteSheetRowNumber++;

    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");
    for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {

      if (null != customerSheet.getRow(j).getCell(1)
              && (customerSheet.getRow(j).getCell(1).getCellType() == CellType.NUMERIC
              || customerSheet.getRow(j).getCell(1).toString() == ""
              || customerSheet.getRow(j).getCell(1).getCellType() == CellType.BLANK)) {

        XSSFRow deletedRow = deletedCustomerSheet.createRow(deleteSheetRowNumber++);

        CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyCellStyle(false);

        deletedRow.copyRowFrom(customerSheet.getRow(j), policy);

        if (customerSheet.getRow(j).getCell(3) == null
                || customerSheet.getRow(j).getCell(3).getCellType() == CellType.BLANK)
          deletedRow
                  .createCell(11)
                  .setCellValue("Reason for deletion : No Email Id Present for this record");
        else
          deletedRow
                  .createCell(11)
                  .setCellValue("Reason for deletion : No Cutomer Id mapping in Export Sheet");

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

    headerRow2.createCell(1).setCellValue("Customer uid ");
    headerRow2.createCell(2).setCellValue("Title");
    headerRow2.createCell(3).setCellValue("FirstName");
    headerRow2.createCell(4).setCellValue("Lastname");
    headerRow2.createCell(5).setCellValue("Line 1");
    headerRow2.createCell(6).setCellValue("Town");
    headerRow2.createCell(8).setCellValue("PostalCode");
    headerRow2.createCell(9).setCellValue("Street Number");
    headerRow2.createCell(10).setCellValue("Street Name");
    headerRow2.createCell(11).setCellValue("Building");
    headerRow2.createCell(12).setCellValue("Mobile Number");
    headerRow2.createCell(13).setCellValue("Country");
    headerRow2.createCell(14).setCellValue("Street Number");
    headerRow2.createCell(15).setCellValue("is billing address");
    headerRow2.createCell(16).setCellValue("is shipping address");
    headerRow2.createCell(17).setCellValue("Collecting point ID");
    headerRow2.createCell(19).setCellValue("Reason");

    deleteSheetRowNumber++;

    for (int j = 0; j <= addressSheet.getLastRowNum(); j++) {

      if (null == addressSheet.getRow(j).getCell(1)
              || addressSheet.getRow(j).getCell(1).getCellType() == CellType.NUMERIC) {

        XSSFRow deletedRow = deletedAddressSheet.createRow(deleteSheetRowNumber++);

        CellCopyPolicy policy = new CellCopyPolicy();
        policy.setCopyCellStyle(false);

        deletedRow.copyRowFrom(addressSheet.getRow(j), policy);

        if (null == addressSheet.getRow(j).getCell(1))
          deletedRow
                  .createCell(19)
                  .setCellValue("Reason for deletion: Customer Id is not present ");
        else
          deletedRow
                  .createCell(19)
                  .setCellValue("Reason for deletion : No Cutomer Id mapping in Export Sheet");

        addressSheet.shiftRows(
                addressSheet.getRow(j).getRowNum() + 1, addressSheet.getLastRowNum() + 1, -1);
        j--;
      }
    }
  }

  private static void exportingFinalAddressWorkbook(XSSFWorkbook addressWorkbook)
          throws IOException {

    FileOutputStream addressFileOutputStream =
            new FileOutputStream(new File("./Final/Address.xlsx"));
    addressWorkbook.write(addressFileOutputStream);

    addressFileOutputStream.close();
  }

  private static void exportingFinalCustomerWorkbook(XSSFWorkbook customerWorkbook)
          throws IOException {
    FileOutputStream customerFileOutputStream =
            new FileOutputStream(new File("./Final/Customer.xlsx"));
    customerWorkbook.write(customerFileOutputStream);
    customerFileOutputStream.close();
  }

  private static List<CSVRecord> importExportCSVFile(Scanner scanner) throws IOException {

    String exportSheetName = null;

    System.out.println("Enter Export Sheet Name with extension:");
    exportSheetName = scanner.nextLine();

    CSVParser exportCSVParser =
            new CSVParser(new FileReader(new File("./Source/" + exportSheetName)), CSVFormat.DEFAULT);
    // exportCSVParser.close();
    return exportCSVParser.getRecords();
  }

  private static XSSFWorkbook importAddressWorkBook(Scanner scanner) throws IOException {
    String addressSheetName = null;
    System.out.println("Enter Address Sheet name with Extension:");
    addressSheetName = scanner.nextLine();

    FileInputStream addressFileInputStream =
            new FileInputStream(new File("./Source/" + addressSheetName));
    // addressFileInputStream.close();
    return new XSSFWorkbook(addressFileInputStream);
  }

  private static XSSFWorkbook importCustomerWorkBook(Scanner scanner) throws IOException {
    String customerSheetName = null;
    System.out.println("Enter Customer Sheet name with extension:");
    customerSheetName = scanner.nextLine();

    FileInputStream customerFileInputStream =
            new FileInputStream(new File("./Source/" + customerSheetName));
    // customerFileInputStream.close();

    return new XSSFWorkbook(customerFileInputStream);
  }

  private static void mappingAddressWorkbook(
          XSSFWorkbook customerWorkbook, XSSFWorkbook addressWorkbook) {

    XSSFSheet addressSheet = addressWorkbook.getSheet("Address");
    XSSFSheet customerSheet = customerWorkbook.getSheet("Customer");

    for (int i = 0; i <= addressSheet.getLastRowNum(); i++) {
      if (null != addressSheet.getRow(i).getCell(1)
              && !addressSheet.getRow(i).getCell(1).toString().equals("")) {

        // System.out.println("Before stripping :" + addressSheet.getRow(i).getCell(1));
        String addressSheetUid =
                StringUtils.stripEnd(addressSheet.getRow(i).getCell(1).toString(), ".0");

        // System.out.println("After stripping :" + addressSheetUid);
        for (int j = 0; j <= customerSheet.getLastRowNum(); j++) {
          //   System.out.println("Before stripping :" + customerSheet.getRow(j).getCell(1));
          if (null != customerSheet.getRow(j).getCell(1)
                  && customerSheet.getRow(j).getCell(1).toString().contains("##")) {

            String[] customerSheetId = customerSheet.getRow(j).getCell(1).toString().split("##");
            String customerUid = customerSheetId[0];

            //   System.out.println("After stripping :" + customerUid);

            if (customerUid.equalsIgnoreCase(addressSheetUid)) {
              String id = customerSheetId[1];
              addressSheet.getRow(i).createCell(1).setCellValue(id);
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
      }

      if (null != customerSheet.getRow(j).getCell(7)
              && customerSheet.getRow(j).getCell(7).toString().equalsIgnoreCase("Registred")) {
        customerSheet.getRow(j).createCell(7).setCellValue("Guest");
      }
    }
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
            /*  System.out.println(
            "Email =" + customerSheet.getRow(i).getCell(3) + " UID =" + record.get(1));*/
            String uid = StringUtils.stripEnd(customerSheet.getRow(i).getCell(1).toString(), ".0");
            // System.out.println("Uid = " + uid);
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
}
