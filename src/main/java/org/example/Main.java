package org.example;

// Import necessary libraries for handling Excel files and Selenium WebDriver
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;
import java.util.Random;

/**
 * This program automates Google searches using keywords from an Excel file.
 * It retrieves the longest and shortest search suggestions and writes them back into the Excel file.
 */
public class Main {

    // Path to the Excel file where search keywords are stored
    private static final String EXCEL_FILE_PATH = ""D:/Projects/GoogleSearchAutomation/Excel.xlsx"";

    // Path to the Chrome WebDriver executable
    private static final String CHROME_DRIVER_PATH = "D:/Projects/GoogleSearchAutomation/chromedriver-win64/chromedriver.exe";

    public static void main(String[] args) throws IOException {
        // Step 1: Set up the Chrome browser
        WebDriver driver = setupWebDriver();

        // Step 2: Read data from the Excel file
        Workbook workbook = readExcelFile();
        FileInputStream fileInputStream = new FileInputStream(new File(EXCEL_FILE_PATH));

        // Step 3: Determine today's day of the week
        DayOfWeek today = LocalDate.now().getDayOfWeek();

        // Get the sheet corresponding to today's day
        Sheet sheet = workbook.getSheet(today.name());

        // If no data is available for today, exit the program
        if (sheet == null) {
            System.out.println("No data for today: " + today.name());
            driver.quit(); // Close the browser
            return;
        }

        // Step 4: Perform Google searches and collect autocomplete suggestions
        performSearchOperation(driver, sheet);

        // Step 5: Write back the results into the Excel file
        writeDataToExcelFile(workbook, fileInputStream);

        // Step 6: Close the browser
        driver.quit();
    }

    /**
     * This method sets up the Chrome WebDriver and applies necessary configurations.
     * @return WebDriver instance for interacting with the browser.
     */
    private static WebDriver setupWebDriver() {
        System.setProperty("webdriver.chrome.driver", CHROME_DRIVER_PATH); // Set the path for ChromeDriver

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");  // Open Chrome in full-screen mode
        options.addArguments("--disable-dev-shm-usage"); // Prevent issues in headless mode
        options.addArguments("--no-sandbox"); // Improve stability in some environments
        options.addArguments("--disable-gpu"); // Disable GPU rendering for better compatibility

        return new ChromeDriver(options); // Create a new Chrome browser instance
    }

    /**
     * Reads the Excel file and returns a Workbook object containing all data.
     * @return Workbook containing keyword data from the Excel file.
     * @throws IOException If the file cannot be read.
     */
    private static Workbook readExcelFile() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File(EXCEL_FILE_PATH));
        return new XSSFWorkbook(fileInputStream); // Return workbook for .xlsx file format
    }

    /**
     * This method performs a Google search for each keyword from the Excel file.
     * It captures the longest and shortest autocomplete suggestions.
     * @param driver The WebDriver instance for browser interaction.
     * @param sheet The Excel sheet containing search keywords.
     */
    private static void performSearchOperation(WebDriver driver, Sheet sheet) {
        Random rand = new Random(); // Used to create random delays to avoid CAPTCHA
        WebDriverWait wait = new WebDriverWait(driver, java.time.Duration.ofSeconds(10)); // Waits for elements to load

        // Loop through each row in the Excel sheet
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;  // Skip the header row

            // Get the keyword from the second column (index 1)
            Cell keywordCell = row.getCell(1);

            // Skip rows with empty keyword cells
            if (keywordCell == null || keywordCell.getStringCellValue().trim().isEmpty()) continue;

            String keyword = keywordCell.getStringCellValue().trim(); // Extract keyword text

            // Open Google homepage
            driver.get("https://www.google.com");

            // Wait for the search box to be clickable
            wait.until(ExpectedConditions.elementToBeClickable(By.name("q")));

            WebElement searchBox = driver.findElement(By.name("q"));
            searchBox.clear(); // Clear any pre-existing text in the search box

            // Simulate typing the keyword character by character (to appear more human-like)
            for (char c : keyword.toCharArray()) {
                searchBox.sendKeys(String.valueOf(c));
                try {
                    Thread.sleep(200 + rand.nextInt(300)); // Random delay between keystrokes (200-500ms)
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                }
            }

            searchBox.sendKeys(Keys.RETURN); // Press Enter to perform the search

            // Wait for autocomplete suggestions to appear
            try {
                wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("ul[role='listbox']")));

                // Get all autocomplete suggestions
                List<WebElement> suggestions = driver.findElements(By.cssSelector("ul[role='listbox'] li"));

                String longest = "", shortest = keyword; // Initialize variables for storing results

                // Iterate through each suggestion and find the longest and shortest ones
                for (WebElement suggestion : suggestions) {
                    String text = suggestion.getText().trim();
                    if (!text.isEmpty()) {
                        if (text.length() > longest.length()) longest = text;
                        if (text.length() < shortest.length()) shortest = text;
                    }
                }

                // Save the longest and shortest suggestions into the Excel sheet
                row.createCell(2).setCellValue(longest);
                row.createCell(3).setCellValue(shortest);

            } catch (Exception e) {
                System.out.println("No suggestions found for keyword: " + keyword);
            }

            // Wait for a random time (3-7 seconds) before processing the next keyword to avoid CAPTCHA
            try {
                Thread.sleep(3000 + rand.nextInt(4000));
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        }
    }

    /**
     * Writes the collected data back to the Excel file and closes the file stream.
     * @param workbook The updated Workbook object with new data.
     * @param fileInputStream The input stream used to read the Excel file.
     */
    private static void writeDataToExcelFile(Workbook workbook, FileInputStream fileInputStream) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(EXCEL_FILE_PATH)) {
            workbook.write(fileOutputStream); // Save the updated data into the file
            workbook.close(); // Close the workbook
            fileInputStream.close(); // Close the file input stream
        } catch (IOException e) {
            System.err.println("Error writing to the Excel file: " + e.getMessage());
        }
    }
}
