import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ExcelToSqlService {

    public static void main(String[] args) throws IOException {

        String filePath = "src/main/resources/Book.xlsx"; // Specify the path to your Excel file
        String tableName = "data"; // Specify the name of your table

        FileInputStream inputStream = new FileInputStream(new File(filePath));
        Workbook workbook = WorkbookFactory.create(inputStream);

        Sheet sheet = workbook.getSheetAt(0); // Assuming you're reading the first sheet

        // Assuming the first row contains column names
        Row headerRow = sheet.getRow(0);

        Cell newInsertColumnHeaderCell = headerRow.createCell(headerRow.getLastCellNum(), CellType.STRING);
        newInsertColumnHeaderCell.setCellValue("Insert Query");
        Cell newSelectColumnHeaderCell = headerRow.createCell(headerRow.getLastCellNum(), CellType.STRING);
        newSelectColumnHeaderCell.setCellValue("Select Query");
        Cell newDeleteColumnHeaderCell = headerRow.createCell(headerRow.getLastCellNum(), CellType.STRING);
        newDeleteColumnHeaderCell.setCellValue("Delete Query");

        List<String> insertQueries = new ArrayList<>();
        List<String> selectQueries = new ArrayList<>();
        List<String> deleteQueries = new ArrayList<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            StringBuilder insertQueryBuilder = new StringBuilder();
            StringBuilder selectQueryBuilder = new StringBuilder();
            StringBuilder deleteQueryBuilder = new StringBuilder();

            // Construct Insert query
            insertQueryBuilder.append("INSERT INTO ").append(tableName).append(" VALUES (");
            // Construct Select query
            selectQueryBuilder.append("SELECT * FROM ").append(tableName).append(" WHERE ");
            // Construct Delete query
            deleteQueryBuilder.append("DELETE FROM ").append(tableName).append(" WHERE ");

            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                if (j > 0) {
                    insertQueryBuilder.append(", ");
                    selectQueryBuilder.append(" AND ");
                }
                String columnName = headerRow.getCell(j).getStringCellValue();
                insertQueryBuilder.append(getCellValueAsString(cell));
                selectQueryBuilder.append(columnName).append(" = ").append(getCellValueAsString(cell));
                deleteQueryBuilder.append(columnName).append(" = ").append(getCellValueAsString(cell));
            }

            insertQueryBuilder.append(");");
            selectQueryBuilder.append(";");
            deleteQueryBuilder.append(";");

            insertQueries.add(insertQueryBuilder.toString());
            selectQueries.add(selectQueryBuilder.toString());
            deleteQueries.add(deleteQueryBuilder.toString());
        }

        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLUE.getIndex());
        style.setFont(font);

        for (int i = 0; i < insertQueries.size(); i++) {
            Row row = sheet.getRow(i + 1);
            Cell insertCell = row.createCell(headerRow.getLastCellNum() - 3, CellType.STRING);
            insertCell.setCellValue(insertQueries.get(i));
            Cell selectCell = row.createCell(headerRow.getLastCellNum() - 2, CellType.STRING);
            selectCell.setCellValue(selectQueries.get(i));
            Cell deleteCell = row.createCell(headerRow.getLastCellNum() - 1, CellType.STRING);
            deleteCell.setCellValue(deleteQueries.get(i));
        }

        // Write back to the Excel file
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);

        workbook.close();
        inputStream.close();
        outputStream.close();
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            return "'" + cell.getStringCellValue() + "'";
        } else if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return "";
        }
    }
}