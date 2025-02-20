import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

public class ExcelHandler extends NanoHTTPD {

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

        if ("read".equals(action)) {
            String result = readExcel(sheetName);
            return newFixedLengthResponse(Response.Status.OK, "text/html", result);
        } else {
            return newFixedLengthResponse(Response.Status.BAD_REQUEST, "text/html", "Invalid action");
        }
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

    private InputStream downloadFile(String fileUrl) throws IOException {
        URL url = new URL(fileUrl);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("GET");
        connection.setConnectTimeout(5000);
        connection.setReadTimeout(5000);
        return connection.getInputStream();
    }
}
