[target](http://www.igapyon.jp/igapyon/diary/2018/ig180418.html) 

2018-04-18 diary: [Java] 大きな Excel ブックを Apache POI で作成
=====================================================================================================
[![いがぴょん画像(小)](http://www.igapyon.jp/igapyon/diary/images/iga200306s.jpg "いがぴょん")](http://www.igapyon.jp/igapyon/diary/memo/memoigapyon.html) [いがぴょん](http://www.igapyon.jp/igapyon/diary/memo/memoigapyon.html)の日記に関連のあるコンテンツ。

## [Java] 大きな Excel ブックを Apache POI で作成

大きな Excel ブックを Apache POI を用いて作成するシンプルなサンプルをメモします。

完全なソースコードは以下にあります。

* [https://github.com/igapy...IWriteExcel](https://github.com/igapyon/simple-POIWriteExcel)

ポイントとなるソースコードは以下。

```java
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

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
```
