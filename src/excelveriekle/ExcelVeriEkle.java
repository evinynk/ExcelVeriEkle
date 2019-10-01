/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excelveriekle;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import javafx.scene.control.Cell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author hdurmaz
 */
public class ExcelVeriEkle {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        
           HSSFWorkbook workbook = new HSSFWorkbook();
           HSSFSheet sheet = workbook.createSheet("Hesaplama");


    Row header = sheet.createRow(0); //Excel içerisindeki satırı temsil eden nesne

    header.createCell(0).setCellValue(" (P)");

    header.createCell(1).setCellValue(" (r)");

    header.createCell(2).setCellValue(" (t)");

    header.createCell(3).setCellValue(" (P r t)");
    
     

    Row dataRow = sheet.createRow(1);

    dataRow.createCell(0).setCellValue(14500d);

    dataRow.createCell(1).setCellValue(9.25);

    dataRow.createCell(2).setCellValue(3d);

    dataRow.createCell(3).setCellFormula("A2*B2*C2"); 

    try {

        FileOutputStream out = 

                new FileOutputStream(new File("C:\\test4.xls"));

        workbook.write(out);

        out.close();

        System.out.println("Excel yazıldı..");

         

    } catch (FileNotFoundException e) {

        e.printStackTrace();

    } catch (IOException e) {

        e.printStackTrace();

    }
    
}
}
