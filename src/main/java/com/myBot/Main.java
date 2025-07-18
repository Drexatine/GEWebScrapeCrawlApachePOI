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
    public static void main(String[] args) throws IOException {

        //How to apply this code to another computer or spreadsheet.
        //On lines 100 and 214, change the file locations to the new location.

        String cellValue;
        String url;

        boolean numeric;

        int rowNumber=0;
        int cellNumber=0;

        int counter=0;

        //start loop here
        do {
            //Limits amount code will run.
            if (counter == 20) {
                break;
            }
            counter = counter+1;

            //Gets the number from the spreadsheet to look up and moves the scanner one down.
            System.out.println("Loading part number.");
            cellValue = getNumber(rowNumber, cellNumber);
            rowNumber = rowNumber + 1;

            //checks if a cell has no value, and if so, then will end the code.
            if (cellValue == null) {
                break;
            }

            //checks if the PN is numeric or not, and refactors it for proper search results if it is.
            numeric = getIsNumeric(cellValue);

            if (numeric) {
                DecimalFormat decimalFormat = new DecimalFormat("0.#####");
                cellValue = decimalFormat.format(Double.valueOf(cellValue));
            }

            System.out.println(cellValue);

            //This will go to the website and get the URL from it, and check if it is a stock URL or not.
            System.out.println("Loading website and scraping information.");
            url = getURL(cellValue);

            if (url == null) {
                System.out.println("Invalid URL or not reachable.");
                System.out.println(" ");
                getAddToCell(url, rowNumber);
                continue;
            }

            System.out.println(url);

            getAddToCell(url, rowNumber);

            System.out.println(" ");

        } while (true);

        System.out.println("It's finally over. Made by Alexander J. Simkins.");

    }

    public static boolean getIsNumeric(String str) {
        try {
            Double.parseDouble(str);
            return true;
        } catch(NumberFormatException e){
            return false;
        }
    }

    public static String getNumber (int rowNumber, int cellNumber) {
        String path = "C:\\Users\\alejo\\Downloads\\apache-poi-src-5.3.0-20240625\\GEHealthcare.xlsx";

        String cellValue = "";

        try (InputStream fileInputStream = Files.newInputStream(Paths.get(path));
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = xssfWorkbook.getSheet("Sheet1");
            if (sheet != null) {
                Row row = sheet.getRow(rowNumber);
                if (row != null) {
                    Cell cell = row.getCell(cellNumber);
                    if (cell != null) {
                        cellValue = switch (cell.getCellType()) {
                            case STRING -> {
                                System.out.println("String value: " + cell.getStringCellValue());
                                yield cell.getStringCellValue();
                            }
                            case NUMERIC -> {
                                System.out.println("Numeric value: " + cell.getNumericCellValue());
                                yield cell.toString();
                            }
                            case BLANK -> {
                                System.out.println("Blank cell");
                                yield null;
                            }
                            default -> cellValue;
                        };
                    } else {
                        System.out.println("Cell is null.");
                        return null;
                    }
                } else {
                    System.out.println("Row is null.");
                    return null;
                }
            } else {
                System.out.println("Sheet is null.");
                return null;
            }

        } catch (IOException e) {
            System.err.println("An I/O error occurred: " + e.getMessage());
            return null;
        }

        return cellValue;
    }

    public static String getURL(String cellValue) throws IOException {
        Document doc;

        Objects object = new Objects();

        // Validate the URL first
        boolean works = getValidURL("https://services.gehealthcare.com/gehcstorefront/p/" + cellValue);

        if (!works) {
            return null;
        }

        //Test to see if it's safe to connect to the site
        try {
            Jsoup.connect("https://services.gehealthcare.com/gehcstorefront/p/" + cellValue).get();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        doc = Jsoup
                .connect("https://services.gehealthcare.com/gehcstorefront/p/" + cellValue)
                .userAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
                .header("Accept-Language", "*")
                .get();

        Elements objectElements = doc.select("div.productDetailsPage");

        for (Element objectElement : objectElements) {
            object.setImage(objectElement.selectFirst("img").attr("src"));
        }

        String url = object.getImage();

        // Check if url is null before using equals method
        if (url == null || url.equals("/gehcstorefront/_ui/desktop/theme-green/images/missing-product-new-300x300.png") || url.equals("/gehcstorefront/_ui/desktop/theme-green/images/missing-product-new-2025-300x300.jpg")) {
            return null;
        }

        //System.out.println(name + "  https://services.gehealthcare.com" + url);
        url = "https://services.gehealthcare.com" + url;

        return url;
    }


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

    public static String getAddToCell (String url, int rowNumber) {

        String path = "C:\\Users\\alejo\\Downloads\\apache-poi-src-5.3.0-20240625\\GEHealthcare.xlsx";

        int rowIndex = rowNumber - 1; // Row, numbers
        int cellIndex = 1; // Cell, letter. Adjust to proper placement when it is working in a new spreadsheet
        String failMessage = "Invalid Image URL or Unreachable Site.";

        try (InputStream fileInputStream = Files.newInputStream(Paths.get(path));
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = xssfWorkbook.getSheetAt(0);

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            Cell cell = row.getCell(cellIndex);
            if (cell == null) {
                cell = row.createCell(cellIndex);
            }

            cell.setCellValue(url);

            if (url == null){
                cell.setCellValue(failMessage);
            }

            try (FileOutputStream fileOutputStream = new FileOutputStream(path)) {
                xssfWorkbook.write(fileOutputStream);
            }

            if (url != null){
                System.out.println("Image URL successfully written to cell " + (char)('A' + cellIndex) + (rowIndex + 1) + ".");
            }

        } catch (IOException e) {
            System.err.println("An I/O error occurred: " + e.getMessage());
        }
        return url;
    }
}