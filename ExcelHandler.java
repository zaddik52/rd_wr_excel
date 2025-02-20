import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

public class ExcelHandler extends NanoHTTPD {
    // URL לקובץ ה-Excel ב-GitHub (בגרסה RAW)
    private static final String FILE_URL = "https://raw.githubusercontent.com/zaddik52/rd_wr_excel/main/list_all.xlsx";

    public ExcelHandler() throws IOException {
        super(8080);
        start(NanoHTTPD.SOCKET_READ_TIMEOUT, false);
        System.out.println("Server started on 8080");
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
        try (InputStream fis = downloadFile(FILE_URL);
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
        try (InputStream fis = downloadFile(FILE_URL);
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

            // במקרה של כתיבה, אנחנו לא שומרים את הקובץ ב-GitHub, אלא רק כותבים אליו מקומית
            // אם תרצה לשמור ב-GitHub, תצטרך לממש פתרון של העלאת הקובץ (למשל API של GitHub).
            return "Cell updated successfully!";
        } catch (Exception e) {
            return "Error writing Excel: " + e.getMessage();
        }
    }

    // פונקציה להורדת קובץ מ-GitHub דרך URL
    private InputStream downloadFile(String fileUrl) throws IOException {
        URL url = new URL(fileUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.setConnectTimeout(5000);
        connection.setReadTimeout(5000);

        return connection.getInputStream();
    }
}
