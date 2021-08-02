package pits;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.StringUtils;
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

      exportingFinalCustomerWorkbook(customerWorkbook);

      exportingFinalAddressWorkbook(addressWorkbook);

      long end = System.currentTimeMillis();

      System.out.println("Time took = " + (end - start));

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static void exportingFinalAddressWorkbook(XSSFWorkbook addressWorkbook) throws IOException {

    FileOutputStream addressFileOutputStream =
            new FileOutputStream(new File("./Final/Address.xlsx"));
    addressWorkbook.write(addressFileOutputStream);

    addressFileOutputStream.close();
  }

  private static void exportingFinalCustomerWorkbook(XSSFWorkbook customerWorkbook) throws IOException {
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
    //customerFileInputStream.close();

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
