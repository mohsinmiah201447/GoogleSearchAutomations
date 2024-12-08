package sampleProject;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;

public class GoogleSearchAutomation {
    public static void main(String[] args) {
        // Set up WebDriver
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\user\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        try {
            // Open the Excel file
            String filePath = "E:\\Assignment Job\\Assessment\\PracticeTask.xlsx";
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Get the current day of the week
            DayOfWeek day = LocalDate.now().getDayOfWeek();
            String sheetName = day.name(); // "MONDAY", "TUESDAY", etc.
            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                System.out.println("Sheet for " + sheetName + " not found in the Excel file.");
                return;
            }

            // Process each keyword in the sheet
            for (Row row : sheet) {
                Cell keywordCell = row.getCell(1); // Assuming keywords are in column B
                if (keywordCell == null || keywordCell.getCellType() != CellType.STRING) continue;

                String keyword = keywordCell.getStringCellValue().trim();
                if (keyword.isEmpty()) continue;

                // Perform Google search
                driver.get("https://www.google.com");
                WebElement searchBox = driver.findElement(By.name("q"));
                searchBox.clear();
                searchBox.sendKeys(keyword);
                searchBox.submit();

                // Extract Google search suggestions
                List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']//li//span"));
                String longest = "", shortest = "";

                for (WebElement suggestion : suggestions) {
                    String text = suggestion.getText();
                    if (!text.isEmpty()) {
                        if (text.length() > longest.length()) longest = text;
                        if (shortest.isEmpty() || text.length() < shortest.length()) shortest = text;
                    }
                }

                // Write results back to the Excel file
                row.createCell(2).setCellValue(longest); // Longest Option in column C
                row.createCell(3).setCellValue(shortest); // Shortest Option in column D
            }

            // Save changes to the file
            fileInputStream.close();
            FileOutputStream outFile = new FileOutputStream(new File(filePath));
            workbook.write(outFile);
            outFile.close();

            System.out.println("Longest and shortest options saved to the sheet: " + sheetName);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit(); // Close the browser
        }
    }
}
