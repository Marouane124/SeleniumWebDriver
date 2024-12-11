package org.example;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.time.Duration;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:\\VariablesEnvironement\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Companies Data");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Nom");
        headerRow.createCell(1).setCellValue("Secteur d'Activité");
        headerRow.createCell(2).setCellValue("Forme Juridique");
        headerRow.createCell(3).setCellValue("Capital");

        int rowNum = 1;

        try {
            // Recuperation du nombre des elements dans la page
            driver.get("https://www.charika.ma");
            performSearch(driver, wait);

            String companyXPath = "//h5[@class='strong text-lowercase truncate']/a[@class='goto-fiche']";

            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(companyXPath)));

            List<WebElement> companyLinks = driver.findElements(By.xpath(companyXPath));
            int totalCompanies = companyLinks.size();
            
            for (int i = 0; i < totalCompanies; i++) {
                try {
                    driver.get("https://www.charika.ma");
                    performSearch(driver, wait);

                    wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath(companyXPath)));
                    WebElement companyLink = wait.until(ExpectedConditions.elementToBeClickable(
                        driver.findElements(By.xpath(companyXPath)).get(i)
                    ));
                    
                    String companyName = companyLink.getText();
                    System.out.println("\nProcessing: " + companyName);
                    companyLink.click();
                    Thread.sleep(2000);

                    String secteurActivite = "Pas disponible";
                    String formeJuridique = "Pas disponible";
                    String capital = "Pas disponible";

                    try {
                        WebElement secteurElement = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[1]/div[2]/span/h2"));
                        if (secteurElement != null && !secteurElement.getText().isEmpty()) {
                            secteurActivite = secteurElement.getText();
                        }
                    } catch (Exception e) {
                        System.out.println("Secteur d'Activité pas trouvé pour: " + companyName);
                    }
                    
                    try {
                        WebElement formeElement = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[4]/div/div[1]/table/tbody/tr[3]/td[2]"));
                        String formeText = formeElement.getText();
                        // Validate that it looks like a legal form (you might want to adjust this validation)
                        if (formeElement != null && !formeText.isEmpty() && !formeText.matches("\\d+.*")) {
                            formeJuridique = formeText;
                        }
                    } catch (Exception e) {
                        System.out.println("Forme Juridique pas trouvé pour: " + companyName);
                    }
                    
                    try {
                        WebElement capitalElement = driver.findElement(
                            By.xpath("//*[@id=\"fiche\"]/div[1]/div[1]/div/div[2]/div[4]/div/div[1]/table/tbody/tr[4]/td[2]"));
                        String capitalText = capitalElement.getText();
                        // Validate that it looks like a capital value (usually contains numbers)
                        if (capitalElement != null && !capitalText.isEmpty() && capitalText.matches(".*\\d+.*")) {
                            capital = capitalText;
                        }
                    } catch (Exception e) {
                        System.out.println("Capital pas trouvé pour: " + companyName);
                    }

                    // Create a new row in Excel and populate it
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(companyName);
                    row.createCell(1).setCellValue(secteurActivite);
                    row.createCell(2).setCellValue(formeJuridique);
                    row.createCell(3).setCellValue(capital);

                    // Print the information for verification
                    System.out.println("Nom: " + companyName);
                    System.out.println("Secteur d'Activité: " + secteurActivite);
                    System.out.println("Forme Juridique: " + formeJuridique);
                    System.out.println("Capital: " + capital);
                    System.out.println("----------------------------------------");

                } catch (Exception e) {
                    System.out.println("Error processing company " + (i + 1) + ": " + e.getMessage());
                    e.printStackTrace();
                    continue;
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream("MarrakechCompanies.xlsx")) {
                workbook.write(outputStream);
            }
            System.out.println("Excel file has been created successfully!");

        } catch (Exception e) {
            System.out.println("Error occurred: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
            driver.quit();
        }
    }

    private static void performSearch(WebDriver driver, WebDriverWait wait) {
        WebElement region = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("//*[@id=\"national\"]/form/div/div[2]/div/div/div/button/div/div/div")));
        region.click();

        WebElement ville = wait.until(ExpectedConditions.elementToBeClickable(
            By.xpath("//*[@id=\"national\"]/form/div/div[2]/div/div/div/div/div[1]/input")));
        ville.sendKeys("Marrakech");
        ville.submit();
    }
}