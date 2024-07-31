package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class ExtractColumnData {
    public static void main(String[] args) throws IOException {
        //Change file path to where it exists in your computer
        FileInputStream fis = new FileInputStream(new File("Excel File copypath"));

        //this is a workbook instance that refers to the .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        //this is the scanner that takes in the column name from user input
        Scanner scanner = new Scanner(System.in);
        System.out.println("Please enter the column name: ");

        //this is the column name you're going to use to extract
        String columnName = scanner.nextLine();
        columnName.stripLeading();
        columnName.stripTrailing();

        //creates a sheet to retrieve the object via column Name
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        //current row
        XSSFRow row = sheet.getRow(0);
        //arrayList to store each row of the column in
        ArrayList<String> excelItems = new ArrayList<>();

        int columnIndex = 0;
        int numberOfRows = sheet.getPhysicalNumberOfRows();

        for (int i =0; i < row.getLastCellNum(); i++) {
            String item = String.valueOf(row.getCell(i));
            if (item.equals(columnName)) {
                columnIndex = i;
            }

        }

        for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
            Row thisRow = CellUtil.getRow(rowIndex, sheet);
            Cell cell = CellUtil.getCell(thisRow, columnIndex);
            excelItems.add(String.valueOf(cell));
        }
        //change file path to where it exists in your computer
        FileWriter writer = new FileWriter("txt file copypath");
        for (String s: excelItems) {
            writer.write(s + System.lineSeparator());
        }
        writer.close();

    }

}
