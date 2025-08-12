package FinalProjectUseCase;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.*;

import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class usecase {

    WebDriver driver;
    WebDriverWait wait;
    String filePathWsr, filePathUpdate, filePathSprint;

    @BeforeClass
    public void OpenApp() throws InterruptedException {
        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.get("http://localhost:5173/");
        
        filePathWsr = System.getProperty("user.dir") + "\\testdata\\Wsr.xlsx";
        filePathUpdate = System.getProperty("user.dir") + "\\testdata\\Update.xlsx";
        filePathSprint = System.getProperty("user.dir") + "\\testdata\\Sprint.xlsx";
    }

    @Test(priority = 1)
    public void CreateProjectWithEmployees() throws IOException, InterruptedException {
        // Navigate to project creation page
        driver.findElement(By.xpath("//*[@id='root']/header/div/button")).click();
        driver.findElement(By.xpath("//a[@href='/project/create']")).click();

        int rows = ExcelUtils.getRowCount(filePathWsr, "Sheet1");
        int employeeIndex = 1;

        for (int i = 1; i <= rows; i++) {
            String projectName = ExcelUtils.getCellData(filePathWsr, "Sheet1", i, 0);
            String emp = ExcelUtils.getCellData(filePathWsr, "Sheet1", i, 1);
            String role = ExcelUtils.getCellData(filePathWsr, "Sheet1", i, 2);
            String expected = ExcelUtils.getCellData(filePathWsr, "Sheet1", i, 3);

            if (!projectName.equals("")) {
                WebElement projectField = driver.findElement(By.xpath("(//input[@type='text'])[1]"));
                projectField.clear();
                projectField.sendKeys(projectName);
                employeeIndex = 1;
            }

            WebElement empBox = driver.findElement(By.xpath("(//input[@role='combobox'])[" + ((employeeIndex - 1) * 2 + 1) + "]"));
            empBox.click();
            empBox.clear();
            empBox.sendKeys(emp);
            driver.findElement(By.xpath("//li[normalize-space()='" + emp + "']")).click();

            WebElement roleBox = driver.findElement(By.xpath("(//input[@role='combobox'])[" + ((employeeIndex - 1) * 2 + 2) + "]"));
            roleBox.click();
            roleBox.clear();
            roleBox.sendKeys(role);
            driver.findElement(By.xpath("//li[normalize-space()='" + role + "']")).click();

            if (i < rows) {
                String nextProjectName = ExcelUtils.getCellData(filePathWsr, "Sheet1", i + 1, 0);
                if (nextProjectName.equals("")) {
                    driver.findElement(By.xpath("//button[@aria-label='Add Employee']")).click();
                    employeeIndex++;
                }
            }
        }

        driver.findElement(By.xpath("//button[normalize-space()='Submit']")).click();

        WebElement popupBtn = driver.findElement(By.xpath("//button[normalize-space()='ok']"));
        String popupText = popupBtn.getText();
        String result = popupText.equalsIgnoreCase("ok") ? "Pass" : "Fail";
        System.out.println("Result: " + result);

        ExcelUtils.setCellData(filePathWsr, "Sheet1", rows, 4, result);

        if (result.equals("Pass")) {
            ExcelUtils.fillGreenColor(filePathWsr, "Sheet1", rows, 4);
        } else {
            ExcelUtils.fillRedColor(filePathWsr, "Sheet1", rows, 4);
        }

        popupBtn.click();
    }

    @Test(priority = 2)
    public void NavToHomePage() {
        WebElement hamburgerMenu = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//*[@id='root']/header/div/button")));
        hamburgerMenu.click();

        WebElement homeLink = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//a[@href='/' and .//span[text()='Home']]")));
        homeLink.click();
    }

    @Test(priority = 3)
    public void ClickUpdate() {
        WebElement updateBtn = wait.until(ExpectedConditions.elementToBeClickable(
                By.xpath("//tbody/tr[1]/td[2]/div[1]/div[1]/button[1]")));
        updateBtn.click();
        System.out.println("Update button clicked");
    }

    @Test(priority = 4)
    public void UpdateEmployeeRoles() throws IOException, InterruptedException {
        // Delete all existing rows
        List<WebElement> deleteIcons = driver.findElements(By.xpath("//button[@aria-label='Delete']//*[name()='svg']"));
        for (WebElement icon : deleteIcons) {
            icon.click();
            Thread.sleep(500);
        }

        int employeeIndex = 1;

        // Fill first 2 rows
        for (int i = 1; i <= 2; i++) {
            String emp = ExcelUtils.getCellData(filePathUpdate, "Sheet1", i, 0);
            String role = ExcelUtils.getCellData(filePathUpdate, "Sheet1", i, 1);

            WebElement empDropdown = driver.findElement(By.xpath("(//button[@title='Open'])[" + ((employeeIndex - 1) * 2 + 1) + "]"));
            empDropdown.click();
            Thread.sleep(500);
            driver.findElement(By.xpath("//li[normalize-space()='" + emp + "']")).click();

            WebElement roleDropdown = driver.findElement(By.xpath("(//button[@title='Open'])[" + ((employeeIndex - 1) * 2 + 2) + "]"));
            roleDropdown.click();
            Thread.sleep(500);
            driver.findElement(By.xpath("//li[normalize-space()='" + role + "']")).click();

            employeeIndex++;

            if (i == 2) {
                driver.findElement(By.xpath("//button[@aria-label='Add Employee']")).click();
                Thread.sleep(1000);
            }
        }

        // Fill third row
        String emp3 = ExcelUtils.getCellData(filePathUpdate, "Sheet1", 3, 0);
        String role3 = ExcelUtils.getCellData(filePathUpdate, "Sheet1", 3, 1);

        WebElement empDropdown3 = driver.findElement(By.xpath("(//button[@title='Open'])[" + ((employeeIndex - 1) * 2 + 1) + "]"));
        empDropdown3.click();
        Thread.sleep(500);
        driver.findElement(By.xpath("//li[normalize-space()='" + emp3 + "']")).click();

        WebElement roleDropdown3 = driver.findElement(By.xpath("(//button[@title='Open'])[" + ((employeeIndex - 1) * 2 + 2) + "]"));
        roleDropdown3.click();
        Thread.sleep(500);
        driver.findElement(By.xpath("//li[normalize-space()='" + role3 + "']")).click();
    }

    @Test(priority = 5)
    public void SubmitUpdate() {
        driver.findElement(By.xpath("//button[normalize-space()='Submit']")).click();
    }

    @Test(priority = 6)
    public void ConfirmPopup() {
        WebElement okBtn = driver.findElement(By.xpath("/html/body/div[2]/div[3]/div/div[2]/button"));
        okBtn.click();
    }
    


    @AfterClass
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }
}
