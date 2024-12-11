package sampleProject;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;

public class GoogleSearchAutomation {
    public static void main(String[] args) {
        // Path to the Excel file
        String excelFilePath = "E:\\Assignment\\Assessment\\Test.xlsx";

        // Set up Selenium WebDriver
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\user\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe");
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-gpu"); // Run without headless mode for better debugging
        WebDriver driver = new ChromeDriver(options);
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

        try {
            // Load the Excel file
            FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Get the current day of the week
            String currentDay = LocalDate.now().getDayOfWeek().getDisplayName(TextStyle.FULL, Locale.ENGLISH);

            // Get the sheet for the current day
            Sheet sheet = workbook.getSheet(currentDay);
            if (sheet == null) {
                System.out.println("No sheet found for " + currentDay);
                return;
            }

            // Iterate over rows starting from row 2 (assuming row 1 is headers)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                if (row != null) {
                    Cell keywordCell = row.getCell(1); // Column B (index 1)
                    if (keywordCell != null) {
                        String keyword = keywordCell.getStringCellValue();

                        // Perform Google search
                        driver.get("https://www.google.com");
                        WebElement searchBox = driver.findElement(By.name("q"));
                        searchBox.sendKeys(keyword);
                        Thread.sleep(2000); // Wait for autocomplete suggestions

                        // Capture autocomplete suggestions
                        List<WebElement> suggestions = driver.findElements(By.cssSelector("ul[role='listbox'] li span"));

                        String longestSuggestion = "";
                        String shortestSuggestion = "";

                        for (WebElement suggestion : suggestions) {
                            String text = suggestion.getText();

                            if (!text.isEmpty()) {
                                if (longestSuggestion.isEmpty() || text.length() > longestSuggestion.length()) {
                                    longestSuggestion = text;
                                }

                                if (shortestSuggestion.isEmpty() || text.length() < shortestSuggestion.length()) {
                                    shortestSuggestion = text;
                                }
                            }
                        }

                        // Write longest and shortest suggestions back to the Excel file
                        Cell longestCell = row.createCell(2, CellType.STRING); // Column C (index 2)
                        Cell shortestCell = row.createCell(3, CellType.STRING); // Column D (index 3)

                        longestCell.setCellValue(longestSuggestion);
                        shortestCell.setCellValue(shortestSuggestion);
                    }
                }
            }

            // Save the updated Excel file
            fileInputStream.close();
            FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();

            System.out.println("Automation completed and Excel file updated successfully for " + currentDay);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit(); // Close the browser
        }
    }
}



