import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Base64;
import java.util.Map;
import org.json.JSONObject;

public class ExcelHandler extends NanoHTTPD {

    private static final String FILE_PATH = "https://github.com/zaddik52/rd_wr_excel/blob/main/list_all.xlsx";
    private static final String GITHUB_REPO = "zaddik52/rd_wr_excel";
    private static final String FILE_NAME = "list_all.xlsx";
    private static final String GITHUB_API_URL = "https://rdwrexcel-production.up.railway.app/list_all.xlsx";
    private static final String GITHUB_TOKEN = System.getenv("GITHUB_TOKEN");  

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
        System.out.println("Request received: " + session.getUri());
        Map<String, String> params = session.getParms();
        String action = params.getOrDefault("action", "read");
        String sheetName = params.getOrDefault("sheet", "Sheet1");

        if ("/read".equals(session.getUri()) && "read".equals(action)) {
            return newFixedLengthResponse(Response.Status.OK, "text/html", readExcel(sheetName));
        }

        if ("/write".equals(session.getUri()) && "write".equals(action)) {
            String cell = params.get("cell");
            String value = params.get("value");
            return newFixedLengthResponse(Response.Status.OK, "text/html", writeExcel(sheetName, cell, value));
        }

        return newFixedLengthResponse(Response.Status.NOT_FOUND, "text/html", "Page not found");
    }

    private String readExcel(String sheetName) {
        try {
            URL url = new URL(FILE_PATH);
            try (InputStream fis = url.openStream();
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
            }
        } catch (Exception e) {
            return "Error reading Excel: " + e.getMessage();
        }
    }

    private String writeExcel(String sheetName, String cellRef, String value) {
        try {
            URL url = new URL(FILE_PATH);
            File tempFile = File.createTempFile("tempExcel", ".xlsx");

            try (InputStream fis = url.openStream();
                 Workbook workbook = new XSSFWorkbook(fis);
                 FileOutputStream fos = new FileOutputStream(tempFile)) {

                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) return "Sheet not found";

                int rowIndex = cellRef.charAt(1) - '1';
                int colIndex = cellRef.charAt(0) - 'A';

                Row row = sheet.getRow(rowIndex);
                if (row == null) row = sheet.createRow(rowIndex);
                Cell cell = row.getCell(colIndex);
                if (cell == null) cell = row.createCell(colIndex);
                cell.setCellValue(value);

                workbook.write(fos);
            }

            return uploadToGitHub(tempFile);
        } catch (Exception e) {
            return "Error writing Excel: " + e.getMessage();
        }
    }

    private String uploadToGitHub(File file) {
        try {
            byte[] fileBytes = new byte[(int) file.length()];
            try (FileInputStream fis = new FileInputStream(file)) {
                fis.read(fileBytes);
            }
            String encodedContent = Base64.getEncoder().encodeToString(fileBytes);
            String sha = getFileSHA();

            JSONObject json = new JSONObject();
            json.put("message", "Updating Excel file");
            json.put("content", encodedContent);
            if (sha != null) {
                json.put("sha", sha);
            }

            HttpURLConnection conn = (HttpURLConnection) new URL(GITHUB_API_URL).openConnection();
            conn.setRequestMethod("PUT");
            conn.setRequestProperty("Authorization", "token " + GITHUB_TOKEN);
            conn.setRequestProperty("Accept", "application/vnd.github.v3+json");
            conn.setDoOutput(true);

            try (OutputStream os = conn.getOutputStream()) {
                os.write(json.toString().getBytes());
            }

            int responseCode = conn.getResponseCode();
            return responseCode == 200 || responseCode == 201 ? "File updated successfully!" : "Failed to update file: HTTP " + responseCode;
        } catch (Exception e) {
            return "Error uploading file to GitHub: " + e.getMessage();
        }
    }

    private String getFileSHA() {
        try {
            HttpURLConnection conn = (HttpURLConnection) new URL(GITHUB_API_URL).openConnection();
            conn.setRequestMethod("GET");
            conn.setRequestProperty("Authorization", "token " + GITHUB_TOKEN);
            conn.setRequestProperty("Accept", "application/vnd.github.v3+json");

            if (conn.getResponseCode() == 200) {
                try (BufferedReader reader = new BufferedReader(new InputStreamReader(conn.getInputStream()))) {
                    String response = reader.lines().reduce("", String::concat);
                    JSONObject json = new JSONObject(response);
                    return json.getString("sha");
                }
            }
        } catch (Exception e) {
            System.err.println("Error fetching file SHA: " + e.getMessage());
        }
        return null;
    }
}
