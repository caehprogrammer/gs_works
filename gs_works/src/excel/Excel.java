/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Carlos Antonio
 */
public class Excel {

    private static final String FILE_NAME = "./xlsx/Libro1.xlsx";
    private static final String FILE_NAME2 = "./xlsx/Libro2.xlsx";

    public void openFile() {
        try (FileInputStream excelFile = new FileInputStream(new File(FILE_NAME))) {
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Cell cell = datatypeSheet.getRow(0).getCell(0);
            System.out.println("" + cell.getCellComment());
            if (cell.getCellComment() != null) {
                cell.removeCellComment();
            }
            this.setComment(cell, "hola2");

            FileOutputStream excelFile2 = new FileOutputStream(new File(FILE_NAME2));
            workbook.write(excelFile2);
            workbook.close();
        } catch (IOException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void setComment(Cell cell, String message) {
        Drawing drawing = cell.getSheet().createDrawingPatriarch();
        CreationHelper factory = cell.getSheet().getWorkbook().getCreationHelper();

        // When the comment box is visible, have it show in a 1x3 space
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 1);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + 1);
        anchor.setDx1(100);
        anchor.setDx2(1000);
        anchor.setDy1(100);
        anchor.setDy2(1000);

        // Create the comment and set the text+author
        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString(message);
        comment.setString(str);
        comment.setAuthor("TURNUS");
        // Assign the comment to the cell
        cell.setCellComment(comment);
    }

    public static void main(String[] args) {
        new Excel().openFile();
    }

}
