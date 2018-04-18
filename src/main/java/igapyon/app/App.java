package igapyon.app;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class App {
    public static void main(String[] args) throws IOException {
        System.out.println("Create large Excel workbook.");

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(10)) {
            for (int indexSheet = 0; indexSheet < 4; indexSheet++) {
                final Sheet sheet = workbook.createSheet();

                for (int indexColumn = 0; indexColumn < 20; indexColumn++) {
                    sheet.setColumnWidth(indexColumn, 256 * 3);
                }

                for (int indexRow = 0; indexRow < 100; indexRow++) {
                    final Row row = sheet.createRow(indexRow);
                    for (int indexColumn = 0; indexColumn < 10; indexColumn++) {
                        final Cell cell = row.createCell(indexColumn);
                        cell.setCellValue("Data of (" + indexColumn + ":" + indexRow + ")");
                    }
                }
            }

            try (BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream("./target/aout.xlsx"))) {
                workbook.write(out);
            }

            workbook.dispose();
        }
    }
}
