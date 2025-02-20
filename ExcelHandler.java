import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.*;

public class ExcelHandler extends NanoHTTPD {
//    private static final String FILE_PATH = "https://rdwrexcel-production.up.railway.app/list_all.xlsx";
//    private static final String FILE_PATH = "C://Temp/chtgpt/list_all.xlsx";
//    private static final String FILE_PATH = "https://github.com/zaddik52/rd_wr_excel/list_all.xlsx";
    private static final String FILE_PATH = "list_all.xlsx";

    public ExcelHandler() throws IOException {
        super(8080);
        start(NanoHTTPD.SOCKET_READ_TIMEOUT, false);
        System.out.println("Server started on http://localhost:8080");
    }

    public static void main(String[] args) {
        try {
            new ExcelHandler();
        } catch (IOException e) {
            System.err.println("Couldn't start server: " + e.getMessage());
        }
    }

    @Override
    public Response serve(IHTTPSession session) {
        Map<String, String> params = session.getParms();
        String action = params.getOrDefault("action", "read");
        String sheetName = params.getOrDefault("sheet", "Sheet1");
        String cell = params.getOrDefault("cell", "A1");
        String value = params.get("value");

        if ("write".equals(action) && value != null) {
            String result = writeExcel(sheetName, cell, value);
            return newFixedLengthResponse(Response.Status.OK, "text/html", result);
        }
        

        String result = readExcel(sheetName);
        return newFixedLengthResponse(Response.Status.OK, "text/html", result);
    }

    private String readExcel(String sheetName) {
        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) return "Sheet not found";

            StringBuilder sb = new StringBuilder("<table border='1'>");
            for (Row row : sheet) {
                sb.append("<tr>");
                for (Cell cell : row) {
                    sb.append("<td>").append(cell.toString()).append("</td>");
                }
                sb.append("</tr>");
            }
            sb.append("</table>");
            return sb.toString();
        } catch (Exception e) {
            return "Error reading Excel: " + e.getMessage();
        }
    }

    private String writeExcel(String sheetName, String cellRef, String value) {
        try (FileInputStream fis = new FileInputStream(FILE_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) return "Sheet not found";

            int rowIndex = cellRef.charAt(1) - '1';
            int colIndex = cellRef.charAt(0) - 'A';

            Row row = sheet.getRow(rowIndex);
            if (row == null) row = sheet.createRow(rowIndex);
            Cell cell = row.getCell(colIndex);
            if (cell == null) cell = row.createCell(colIndex);
            cell.setCellValue(value);

            try (FileOutputStream fos = new FileOutputStream(FILE_PATH)) {
                workbook.write(fos);
            }

            return "Cell updated successfully!";
        } catch (Exception e) {
            return "Error writing Excel: " + e.getMessage();
        }
    }
}