package org.example;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.IOException;

import static org.apache.poi.ss.usermodel.CellType.*;

public class Main {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("C:/Users/SerebrovEV/IdeaProjects/Task35/read.xls");
        Workbook wb = new HSSFWorkbook(fis);
        String element = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
        System.out.println(element);

      //  String element2 = wb.getSheetAt(0).getRow(0).getCell(1).getStringCellValue();
        // Через стринг при других значениях ячейки будет ошибка, для этого используем обработчик типа данных в ячейке
        System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(1)));

        for (Row row : wb.getSheetAt(0)){
            for (Cell cell : row){
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");
                System.out.println(getCellText(cell));
            }
        }
        fis.close();
    }

    public static String getCellText(Cell cell){
        String result = "";
        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case FORMULA:
                result = cell.getCellFormula();
                break;
            case BLANK:
                break;
            default:
                result = "";
        }
        return result;
    }

}