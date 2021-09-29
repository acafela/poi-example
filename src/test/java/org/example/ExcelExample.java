package org.example;

import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelExample {

    @Test
    public void readCellTest() {
        ClassLoader classLoader = getClass().getClassLoader();
        try (InputStream xlsxIs = classLoader.getResourceAsStream("sample/sample.xlsx")) {
            assertNotNull(xlsxIs, "xlsxIs");
            Workbook wb = WorkbookFactory.create(xlsxIs);
            Sheet sheet = wb.getSheetAt(0);
            printRows(sheet);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void printRows(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = sheet.getFirstRowNum(); i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            assertNotNull(row, "row");
            printColumn(row);
        }
    }

    private void printColumn(Row row) {
        int lastCellNum = row.getLastCellNum();
        for (int i = row.getFirstCellNum(); i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            assertNotNull(cell, "cell");
            System.out.print(cell.getStringCellValue() + "\t|\t");
        }
        System.out.println();
    }

    @Test
    public void readAllSheetNamesTest() {
        ClassLoader classLoader = getClass().getClassLoader();
        try (InputStream xlsxIs = classLoader.getResourceAsStream("sample/sample.xlsx")) {
            assertNotNull(xlsxIs, "xlsxIs");
            Workbook wb = WorkbookFactory.create(xlsxIs);
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                System.out.print(wb.getSheetName(i) + ", ");
            }
            System.out.println();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
