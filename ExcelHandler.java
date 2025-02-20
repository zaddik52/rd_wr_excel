import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

public class ExcelHandler extends NanoHTTPD {

    private static final String FILE_PATH = "https://raw.githubusercontent.com/zaddik52/rd_wr_excel/main/list_all.xlsx";

    public ExcelHandler() throws IOException {
        super(8080);  // אם זה ב-RAILWAY, נשתמש ב-Port אחר אם יש צורך
        start(NanoHTTPD.SOCKET_READ_TIMEOUT, false);
        System.out.println("Server started on port 8080");
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

        if ("/read".equals(session.getUri())) {
            if ("read".equals(action)) {
                String result = readExcel(sheetName);
                return newFixedLengthResponse(Response.Status.OK, "text/html", result);
            }
        }

        if ("/write".equals(session.getUri())) {
            if ("write".equals(action)) {
                String cell = params.get("cell");
                String value = params.get("value");
                String result = writeExcel(sheetName, cell, value);
                return newFixedLengthResponse(Response.Status.OK, "text/html", result);
            }
        }

        return newFixedLengthResponse(Response.Status.NOT_FOUND, "text/html", "Page not found");
    }

    private String readExcel(String sheetName) {
        try (InputStream fis = new FileInputStream(FILE_PATH);
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
        try (InputStream fis = new FileInputStream(FILE_PATH);
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

            // לשמירה של הקובץ, עליך לשמור אותו מחדש ב-RAILWAY
            try (FileOutputStream fileOut = new FileOutputStream(FILE_PATH)) {
                workbook.write(fileOut);
            }

            return "Cell updated successfully!";
        } catch (Exception e) {
            return "Error writing Excel: " + e.getMessage();
        }
    }
}
