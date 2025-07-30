package com.myBot;

import org.jsoup.*;
import org.jsoup.nodes.*;
import org.jsoup.select.*;

import java.io.FileOutputStream;
import java.io.IOException;

import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URI;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        String path = "C:\\Users\\alejo\\Downloads\\apache-poi-src-5.3.0-20240625\\GEHealthcare.xlsx";
        String sheetName = "Sheet1"; // or just use getSheetAt(0)

        try (InputStream inputStream = Files.newInputStream(Paths.get(path));
             XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("Sheet not found!");
                return;
            }

            int rowNumber = 0;
            int cellNumber = 0;
            int counter = 0;

            while (true) {
                if (counter == 30) break;
                counter++;

                System.out.println("Loading part number.");
                String cellValue = getNumber(sheet, rowNumber, cellNumber);
                rowNumber++;

                if (cellValue == null) break;

                boolean numeric = getIsNumeric(cellValue);
                if (numeric) {
                    DecimalFormat decimalFormat = new DecimalFormat("0.#####");
                    cellValue = decimalFormat.format(Double.parseDouble(cellValue));
                }
                System.out.println(cellValue);

                System.out.println("Loading website and scraping information.");
                String url = getURL(cellValue);

                if (url == null) {
                    System.out.println("Invalid URL or not reachable.\n");
                    addToCell(sheet, rowNumber - 1, 1, null);  // use row index directly
                    continue;
                }

                System.out.println(url);
                addToCell(sheet, rowNumber - 1, 1, url);
                System.out.println(" ");
            }

            // After processing all rows, write back to file
            try (FileOutputStream fileOut = new FileOutputStream(path)) {
                workbook.write(fileOut);
            }

            System.out.println("Code has ended after " + counter + " times looped.");

        } catch (IOException e) {
            System.err.println("Failed to open workbook: " + e.getMessage());
        }
    }

    /**
     * This checks if a part number is numeric or not.
     * @param str the cell value
     * @return true or false
     */
    public static boolean getIsNumeric(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch(NumberFormatException e){
            return false;
        }
    }

    /**
     * This connects to the Spreadsheet and collects the value of a cell
     * @param rowNumber What row to look in
     * @param cellNumber What cell to look in
     * @return The value of the cell
     */
    public static String getNumber(Sheet sheet, int rowNumber, int cellNumber) {
        Row row = sheet.getRow(rowNumber);
        if (row == null) return null;

        Cell cell = row.getCell(cellNumber);
        if (cell == null) return null;

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BLANK -> null;
            default -> cell.toString();
        };
    }

    /**
     * This goes to the parts GE Webshop page and collects the image URl
     * @return The Images Url, or null
     * @throws IOException Protects the code incase it can't connect
     */
    public static String getURL(String partNumber) {
        String baseUrl = "https://services.gehealthcare.com/gehcstorefront/p/";
        String fullUrl = baseUrl + partNumber;

        try {
            // Make a single connection and parse HTML
            Document doc = Jsoup.connect(fullUrl)
                    .userAgent("Mozilla/5.0")
                    .header("Accept-Language", "*")
                    .timeout(50000) // optional timeout
                    .get();

            Elements productDetails = doc.select("div.productDetailsPage img");

            if (productDetails.isEmpty()) {
                System.out.println("No image found.");
                return null;
            }

            String src = productDetails.first().attr("src");
            if (src == null || src.contains("missing")) {
                System.out.println("Stock or missing image.");
                return null;
            }

            return "https://services.gehealthcare.com" + src;

        } catch (IOException e) {
            System.err.println("Failed to fetch page: " + e.getMessage());
            return null;
        }
    }


    /**
     * This test if the site exist or not
     * @param urlString The site to go to
     * @return if it connected or not
     */
    private static boolean getValidURL(String urlString) {
        try {
            URI uri = new URI(urlString);
            URL url = uri.toURL();

            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("HEAD");

            int responseCode = connection.getResponseCode();

            return (200 <= responseCode && responseCode <= 399);

        } catch (Exception e) {
            return false;
        }
    }

    /**
     * This will add the image URL or a fail message to the cell of the part
     * @param url The parts image URL
     */
    public static void addToCell(Sheet sheet, int rowIndex, int cellIndex, String url) {
        String failMessage = "Invalid Image URL or Unreachable Site.";

        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        Cell cell = row.getCell(cellIndex);
        if (cell == null) cell = row.createCell(cellIndex);

        cell.setCellValue(url != null ? url : failMessage);

        if (url != null) {
            System.out.println("Image URL successfully written to cell " +
                    (char) ('A' + cellIndex) + (rowIndex + 1) + ".");
        }
    }

    public class partResult {
        int rowNumber;
        String partNumber;
        String imageURL;

        public partResult(int rowNumber, String partNumber) {
            this.rowNumber = rowNumber;
            this.partNumber = partNumber;
        }

        public void setImageURL(String imageURL) {
            this.imageURL=imageURL;
        }
    }
}