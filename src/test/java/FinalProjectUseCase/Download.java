package FinalProjectUseCase;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.*;

import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class Download {

    WebDriver driver;
    String filePath;

    @BeforeClass
    public void openApplication() {
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.get("http://localhost:5173/");
        driver.manage().window().maximize();
        filePath = System.getProperty("user.dir") + "\\testdata\\Sprint.xlsx";
    }

    @Test(priority = 1)
    public void fillSprintForm() throws IOException, InterruptedException {
        int row = 1;

        // Navigate to Sprint Page
        driver.findElement(By.xpath("//*[@id='root']/header/div/button")).click();
        driver.findElement(By.xpath("//a[.//span[text()='Sprint']]")).click();

        // Select Project Name
        List<WebElement> dropdownIcons = driver.findElements(By.cssSelector("svg[data-testid='ArrowDropDownIcon']"));
        dropdownIcons.get(0).click();
        Thread.sleep(1000);
        List<WebElement> projectOptions = driver.findElements(By.cssSelector("ul[role='listbox'] li"));
        projectOptions.get(0).click();

        // Fill input fields
        driver.findElement(By.name("sprintNo")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 1));
        driver.findElement(By.name("sprintGoal")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 2));
        driver.findElement(By.name("startDate")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 3));
        driver.findElement(By.name("endDate")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 4));
        driver.findElement(By.name("accomplishment")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 5));
        driver.findElement(By.name("risk")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 6));
        driver.findElement(By.name("committedStoryPoints")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 7));
        driver.findElement(By.name("deliveredStoryPoints")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 8));
        driver.findElement(By.name("velocity")).sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 9));

        // Select Status
        dropdownIcons.get(1).click();
        Thread.sleep(1000);
        List<WebElement> statusOptions = driver.findElements(By.cssSelector("ul[role='listbox'] li"));
        statusOptions.get(0).click();

        // Capacity
        WebElement capacity = driver.findElement(By.cssSelector("input[type='number'].MuiInputBase-inputSizeSmall"));
        capacity.sendKeys(ExcelUtils.getCellData(filePath, "Sheet1", row, 11));
        Thread.sleep(500);

        // Submit
        WebElement submitButton = driver.findElement(By.xpath("//button[normalize-space()='Submit']"));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", submitButton);
        submitButton.click();
        Thread.sleep(2000);
    }

    @Test(priority = 2)
    public void confirmSprintPopup() {
        WebElement okBtn = driver.findElement(By.xpath("/html/body/div[2]/div[3]/div/div[2]/button"));
        okBtn.click();
    }

    @Test(priority = 3)
    public void downloadWeeklyReport() throws IOException, InterruptedException {
        // Navigate to Download Page
        driver.findElement(By.xpath("//*[@id='root']/header/div/button")).click();
        driver.findElement(By.xpath("//a[.//span[text()='Download']]")).click();

        List<WebElement> inputFields = driver.findElements(By.xpath("//input[@role='combobox']"));

        WebElement projectInput = inputFields.get(0);
        projectInput.click();
        projectInput.clear();
        projectInput.sendKeys("ecommerce");
        driver.findElement(By.xpath("//li[normalize-space()='ecommerce']")).click();

        WebElement sprintInput = inputFields.get(1);
        sprintInput.click();
        sprintInput.clear();
        sprintInput.sendKeys("1");
        driver.findElement(By.xpath("//li[normalize-space()='1']")).click();

        driver.findElement(By.xpath("//button[normalize-space()='Export Excel']")).click();
        Thread.sleep(2000);
    }

    @AfterClass
    public void closeBrowser() {
        driver.quit();
    }
}
