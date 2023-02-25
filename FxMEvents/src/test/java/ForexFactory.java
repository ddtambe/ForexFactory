import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.time.Duration;
import java.util.List;
import java.util.Objects;

public class ForexFactory {
    ChromeOptions options;

    WebDriver driver;
    WebDriverWait wait;
    @BeforeClass
    public void setup(){
        System.setProperty("Webdriver.chrome.driver", "chromedriver_win32/chromedriver.exe");
        options = new ChromeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--disable-notifications");
        options.addArguments("--disable-infobars");
        options.addArguments("--disable-popup-blocking");

        driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.forexfactory.com/calendar?day=jan2.2007");
        wait = new WebDriverWait(driver, Duration.ofSeconds(30));
    }


    @Test
    public void open() {
//        WebElement yearsBackButton = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("td[class='calendarmini__headernav']>a[class='calJump year back']")));
//        WebElement currentYear = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("td[class='calendarmini__headernav calendarmini__headernav--current current']")));
//        JavascriptExecutor j = (JavascriptExecutor) driver;

        WebElement date = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("td[class='calendar__cell calendar__date date']>span>span")));
        WebElement day = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("td[class='calendar__cell calendar__date date']>span")));

        List<WebElement> blocks = driver.findElements(By.cssSelector("[class*='calendar__row calendar_row calendar__row--grey']"));

        System.out.println("B size: "+blocks.size());
        int i=1;
        for(WebElement b : blocks){
            System.out.println("B: "+b+(i++));
            System.out.println("Date: "+date.getText());
            System.out.println("Day: "+day.getText());
            System.out.println();
        }
    }
    @AfterClass
    public void closeDriver(){
//        driver.quit();
    }
}
