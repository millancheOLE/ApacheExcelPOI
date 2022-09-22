package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("Spisak.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
            XSSFRow row = sheet.getRow(i);

            String kolonaA = "";
            String kolonaB = "";

            for (int j = 0; j < 1; j++) {
                XSSFCell cell = row.getCell(j);
                kolonaA = cell.getStringCellValue();
            }
            for (int k = 1; k < 2; k++) {
                XSSFCell cell = row.getCell(k);
                kolonaB = cell.getStringCellValue();
            }

            if (kolonaA.equals(kolonaB)) {
                for (int j = 2; j < 3; j++) {
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue("Invalid");
                    backgroundColorGreen(cell);
                }
            } else {

                for (int j = 2; j < 3; j++) {
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue("Valid");
                    backgroundColorOrange(cell);
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream("Spisak.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    public static void backgroundColorGreen(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
    }

    public static void backgroundColorOrange(Cell cell) {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
    }
}