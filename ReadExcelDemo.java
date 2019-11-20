import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class ReadExcelDemo {
    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream(new File("/Users/narendran/Desktop/excelfile/report.xls"));

            HSSFWorkbook workbook = new HSSFWorkbook(file);

            HSSFSheet sheet = workbook.getSheetAt(0);

            int row = findRow(sheet, args[0]);
            System.out.print("last name :: " + sheet.getRow(row).getCell(1));

            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static int findRow(HSSFSheet sheet, String cellContent) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        if (cell.getStringCellValue().equals(cellContent)) {
                            return cell.getRowIndex();
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if (String.valueOf(cell.getNumericCellValue()).equals(cellContent)) {
                            return cell.getRowIndex();
                        }
                        break;
                    default:
                        break;
                }
            }
        }
        return 0;
    }
}
