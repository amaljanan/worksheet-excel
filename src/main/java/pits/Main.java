package pits;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;


public class Main {

  public static void main(String[] args) {

    long start = System.currentTimeMillis();

    try {
      FileInputStream customerFileInputStream =
          new FileInputStream(
              "/Users/pituser/Documents/Waltec Project/20210608_PH_FR_Template_Customer.xlsx");

      CSVParser exportCSVParser =
          new CSVParser(
              new FileReader("/Users/pituser/Documents/Waltec Project/Export.csv"),
              CSVFormat.DEFAULT);
      List<CSVRecord> list = exportCSVParser.getRecords();
      exportCSVParser.close();

      XSSFWorkbook customerWorkbook = new XSSFWorkbook(customerFileInputStream);

      mappingCustomerWorkbook(customerWorkbook, list);

      FileInputStream addressFileInputStream =
          new FileInputStream(
              "/Users/pituser/Documents/Waltec Project/20210608_ PH_FR_Template_Address.xlsx");

      XSSFWorkbook addressWorkbook = new XSSFWorkbook(addressFileInputStream);

      mappingAddressWorkbook(customerWorkbook, addressWorkbook);

      FileOutputStream customerFileOutputStream =
          new FileOutputStream(
              "/Users/pituser/Documents/Waltec Project/Result/Customer.xlsx");
      customerWorkbook.write(customerFileOutputStream);

      FileOutputStream addressFileOutputStream =
          new FileOutputStream(
              "/Users/pituser/Documents/Waltec Project/Result/Address.xlsx");
      addressWorkbook.write(addressFileOutputStream);

      addressFileOutputStream.close();
      customerFileOutputStream.close();

      customerFileInputStream.close();

      long end = System.currentTimeMillis();

      System.out.println("Time took = "+(end-start));

    } catch (Exception e) {
      e.printStackTrace();
    }
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
