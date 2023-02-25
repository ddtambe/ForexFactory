package org.example;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.List;

public class ForexEvents {
    public static void main(String[] args) {

        WebDriver driver;
        WebDriverWait wait;
        ChromeOptions options;
        options = new ChromeOptions();
        driver = new ChromeDriver(options);
        driver.manage().deleteAllCookies();
        wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        options.addArguments("--start-maximized");
        options.addArguments("--disable-notifications");
        options.addArguments("--disable-infobars");
        options.addArguments("--disable-popup-blocking");
        options.addArguments("--ignore-certificate-errors");
        options.setAcceptInsecureCerts(true);
        DesiredCapabilities capabilities = new DesiredCapabilities();
        capabilities.setCapability(CapabilityType.PROXY, new Proxy().setHttpProxy("my-proxy-server:8080"));
        options.merge(capabilities);
        driver.get("https://www.forexfactory.com/calendar?day=jan4.2007");

        WebElement date = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("tbody td[class='calendar__cell calendar__date date']>span>span")));
        WebElement day = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("tbody td[class='calendar__cell calendar__date date']>span")));
        List<WebElement>  time = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__time time']")));
        List<WebElement>  currency = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__currency currency ']")));
        List<WebElement>  eventName = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody span.calendar__event-title")));
        List<WebElement>  actualNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__actual actual']")));
        List<WebElement>  forcastNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__forecast forecast']")));
        List<WebElement>  preivousNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__previous previous']")));

        List<String> list = new ArrayList<>();
        System.out.println("Sizes: "+ time.size()+" "+currency.size()+" "+eventName.size()+" "+actualNumbers.size()+" "+forcastNumbers.size()+" "+preivousNumbers.size());

        for(int i=0; i<time.size(); i++){
            list.add(date.getText()+" "+time.get(i).getText()+" "+currency.get(i).getText()
                    +" "+eventName.get(i).getText()+" "+actualNumbers.get(i).getText()
                    +" "+forcastNumbers.get(i).getText()+" "+preivousNumbers.get(i).getText());
        }
        for (String s : list){
            System.out.println(s);
        }
        driver.quit();
    }
}
