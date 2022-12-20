//package org.example;
//
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.ss.usermodel.Cell;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.IOException;
//
//public class Example {
//
//    public void ReadData() throws IOException {
//
//        // импортируйте лист Excel из директории web диска c.
//        //QurentSourceFile – это имя нужного файла excel
//
//        File sourceFile = new File("C:\\web\\QurentSourceFile.xls");
//
//// На этом шаге мы загружаем файл. Мы используем FileInputStream для чтения из
//// файла excel. В случае если вы хотите проводить запись в файл -
//// вам следует использовать FileOutputStream. Путь к файлу передается в качестве
//// аргумента FileInputStream
//
//        FileInputStream fileInput = new FileInputStream(sourceFile);
//
//        // На этом шаге мы загружаем рабочую книгу excel с помощью HSSFWorkbook,
//        // в который мы передаем fileInput в качестве аргумента
//
//        HSSFWorkbook book = new HSSFWorkbook(fileInput);
//
//        // На этом шаге мы загружаем конкретный лист excel, на котором хранятся данные.
//
//        qurentSheet = book.getSheetAt(0);
//
//        for (int i = 1; i <= qurentSheet.getLastRowNum(); i++) {
//            // Import data for Email.
//            qurentCell = qurentSheet.getRow(i).getCell(1);
//            qurentCell.setCellType(Cell.CELL_TYPE_STRING);
//            driver.findElement(By.id("email")).sendKeys(qurenrCell.getStringCellValue());
//
//            // Импортируем данные из ячеек с паролями.
//
//            qurentCell = qurentSheet.getRow(i).getCell(2);
//            qurentCell.setCellType(Cell.CELL_TYPE_STRING);
//            driver.findElement(By.id("password")).sendKeys(qurenrCell.getStringCellValue());
//
//
//        }
//
//    }
//}
