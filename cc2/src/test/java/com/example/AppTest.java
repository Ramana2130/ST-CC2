package com.example;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;

/**
 * Unit test for simple App.
 */
public class AppTest {
    public WebDriver driver;
    public XSSFWorkbook workbook;
    public ExtentReports extentReports;
    public String toSearch;
    public Actions action;
    public JavascriptExecutor js;
    public ExtentTest test1, test2, test3;

    @BeforeTest
    public void setup() throws IOException {
        System.setProperty("webdriver.chrome.driver", "C:\\ProgramFiles\\Google\\Chrome\\Application");
        driver = new ChromeDriver();
        action = new Actions(driver);
        js = (JavascriptExecutor) driver;
        extentReports = new ExtentReports();
        ExtentSparkReporter spark = new ExtentSparkReporter(
                "C:\\Users\\raman\\OneDrive\\Desktop\\SoftwareTesting\\cc2\\src\\report\\report.html");
        extentReports.attachReporter(spark);
        spark.config().setDocumentTitle("Report Title");
        spark.config().setTheme(Theme.DARK);

        workbook = new XSSFWorkbook("C:\\Users\\raman\\OneDrive\\Desktop\\SoftwareTesting\\cc2\\src\\Excel\\Data.xlsx");
        XSSFSheet sheet = workbook.getSheet("Sheet1");
        toSearch = sheet.getRow(1).getCell(0).getStringCellValue();
    }

    @Test
    public void SearchTest() throws Exception {
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.barnesandnoble.com");
        Thread.sleep(3000);

        driver.findElement(By.xpath("//*[@id='rhf_header_element']/nav/div/div[3]/form/div/div[1]/a")).click();
        Thread.sleep(3000);

        driver.findElement(By.xpath("//*[@id='rhf_header_element']/nav/div/div[3]/form/div/div[1]/div/a[2]")).click();
        Thread.sleep(3000);

        driver.findElement(By.xpath("//*[@id='rhf_header_element']/nav/div/div[3]/form/div/div[2]/div/input[1]"))
                .sendKeys(toSearch);
        Thread.sleep(2000);

        driver.findElement(By.xpath("//*[@id='rhf_header_element']/nav/div/div[3]/form/div/span/button")).click();
        Thread.sleep(2000);

        String searchResult = driver
                .findElement(By.xpath("//*[@id='searchGrid']/div/section[1]/section[1]/div/div[1]/div[1]/h1/span"))
                .getText();

        System.out.println(searchResult);
        test1 = extentReports.createTest(searchResult);
        if (searchResult.equals(toSearch)) {
            test1.log(Status.PASS, "Search Results contain the keyword Chetan Bhagat");
        } else {
            test1.log(Status.FAIL, "Search Results does not contain the keyword Chetan Bhagat");

        }
        Thread.sleep(3000);

    }

    @Test
    public void testCase2() throws InterruptedException {
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.barnesandnoble.com");
        Thread.sleep(3000);
        Actions action = new Actions(driver);

        action.moveToElement(driver.findElement(By.xpath("//*[@id=\"rhfCategoryFlyout_Audiobooks\"]"))).perform();
        Thread.sleep(2000);

        driver.findElement(
                By.xpath("//*[@id='navbarSupportedContent']/div/ul/li[5]/div/div/div[1]/div/div[2]/div[1]/dd/a[1]"))
                .click();
        Thread.sleep(2000);
        
        js = (JavascriptExecutor) driver;
        js.executeScript("window.scrollBy(0,1000)", "");
        Thread.sleep(4000);

        driver.findElement(By.xpath("//*[@id=\"addToBagForm_2940159543998\"]/input[11]")).click();
        Thread.sleep(2000);

    }

    @AfterTest
    public void afterTest() {
        extentReports.flush();
        // driver.quit();
    }

}
