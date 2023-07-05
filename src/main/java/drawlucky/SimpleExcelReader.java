package drawlucky;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Random;

public class SimpleExcelReader {

//             Declare main method to run application
    public static void main(String[] args) {
//            Use a try-catch block to catch possible exceptions when working with Excel files.
        try {
//            Creates a FileInputStream object to read data from the file user.xlsx. This file is saved at the path  on your computer.
            FileInputStream excelFile = new FileInputStream("C:\\Users\\ASUS\\Documents\\user.xlsx");
//            Creates a Workbook object to access the contents of the Excel file. Use the XSSFWorkbook class to support the xlsx format.
            Workbook workbook = new XSSFWorkbook(excelFile);
//            Creates a Sheet object to access a sheet in an Excel file. Use the getSheetAt(0) method to get the first sheet in the file.
            Sheet datatypeSheet = workbook.getSheetAt(0);
//            Creates an Iterator<Row> object to cycle through the rows in the sheet. Use sheet's iterator() method to get the iterator.
            Iterator<Row> iterator = datatypeSheet.iterator();
//            Use a while loop to check if there are any rows left in the sheet. If yes, execute the commands inside the loop.
            while (iterator.hasNext()) {
//                Tạo một đối tượng Row để lưu trữ hàng hiện tại trong sheet. Sử dụng phương thức next() của iterator để lấy hàng tiếp theo.
                Row currentRow = iterator.next();
//                Creates an Iterator<Cell> object to cycle through the cells in the row. Use the row's iterator() method to get the iterator.
                Iterator<Cell> cellIterator = currentRow.iterator();
//                Use a while loop to check if there are any cells left in the row. If yes, execute the commands inside the loop.
                while (cellIterator.hasNext()) {
//                    Creates a Cell object to store the current cell in the row. Use the iterator's next() method to get the next cell.
                    Cell currentCell = cellIterator.next();
                    //getCellTypeEnum shown as deprecated for version 3.15
                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
//                    Check the data type of the current cell. If it is a string type, use the getStringCellValue() method
//                    to get the string value of the cell and print it to the screen with the "–" character.
                    if (currentCell.getCellType() == CellType.STRING) {
                        System.out.print(currentCell.getStringCellValue() + "_");
//                        If it is a numeric type, use the getNumericCellValue() method to get the numeric value of the cell and print it to the screen with the "–" character.
                    } else if (currentCell.getCellType() == CellType.NUMERIC) {
                        System.out.print(currentCell.getNumericCellValue() + "_");
                    }
                }
//                After traversing all the cells in the row, print a newline character to start a new row.
                System.out.println();
            }
//            Catch a FileNotFoundException if the user.xlsx file is not found and print an error message.
        } catch (FileNotFoundException e) {
            e.printStackTrace();
//            Catch an IOException if there is an error reading or writing data from the Excel file and print the error message.
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}