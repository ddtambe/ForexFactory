package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.*;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.List;

/*
 * Change link
 * Change Days in for loop Line No 82
 * Change year at line no. 96
 * Add New Sheet in Excel
 * Change Sheet Name
 */
public class ForexEvents {
    public static void main(String[] args) throws InterruptedException {
        WebDriver driver;
        WebDriverWait wait;
        ChromeOptions options;
        options = new ChromeOptions();

        options.addArguments("--start-maximized");
        options.addArguments("--disable-notifications");
        options.addArguments("--disable-infobars");
        options.addArguments("--disable-popup-blocking");
        options.addArguments("--ignore-certificate-errors");
        options.setAcceptInsecureCerts(true);

        DesiredCapabilities capabilities = new DesiredCapabilities();
        capabilities.setCapability(CapabilityType.PROXY, new Proxy().setHttpProxy("my-proxy-server:8080"));
        options.merge(capabilities);
        driver = new ChromeDriver(options);
        driver.manage().deleteAllCookies();
        wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        Queue<String[]> q = new LinkedList<>();
        q.offer(new String[]{"https://www.forexfactory.com/calendar?day=mar", "31",".2023"});

        while (!(q.isEmpty())) {
            String[] vals = q.poll();
            for (int dayNumber = 1; dayNumber<= Integer.parseInt(vals[1]) ; dayNumber++) {
                driver.manage().deleteAllCookies();
                driver.get(vals[0] + dayNumber + vals[2]);
                WebElement date = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("tbody td[class='calendar__cell calendar__date date']>span>span")));
                WebElement day = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("tbody td[class='calendar__cell calendar__date date']>span")));
                List<WebElement> time = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__time time']")));
                List<WebElement> currency = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__currency currency ']")));
                List<WebElement> eventName = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody span.calendar__event-title")));
                List<WebElement> actualNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__actual actual']")));
                List<WebElement> forcastNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__forecast forecast']")));
                List<WebElement> preivousNumbers = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.cssSelector("tbody td[class='calendar__cell calendar__previous previous']")));

                List<ForexCalender> list = new ArrayList<>();
//            System.out.println("Sizes: " + time.size() + " " + currency.size() + " " + eventName.size() + " " + actualNumbers.size() + " " + forcastNumbers.size() + " " + preivousNumbers.size());

                for (int i = 0; i < time.size(); i++) {
                    list.add(new ForexCalender(date.getText() + " " + vals[2].substring(1), time.get(i).getText(), currency.get(i).getText(),
                            eventName.get(i).getText(), actualNumbers.get(i).getText(), forcastNumbers.get(i).getText(), preivousNumbers.get(i).getText()));
                }

//            for(int i=0; i<time.size(); i++)
//                System.out.println(list.get(i));
                System.out.println("Date: " + date.getText() + " | Size: " + list.size());

                writeExcel(list, vals[2]);
                list.clear();
                System.out.println("Excel Writing Done");
                System.out.println();
                driver.manage().deleteAllCookies();
                driver.switchTo().newWindow(WindowType.WINDOW);

                Set<String> handles = driver.getWindowHandles();
//                System.out.println(handles);
                for (String handle : handles) {
                    if (!handle.equals(driver.getWindowHandle())) {
                        driver.switchTo().window(handle);
                    }
                }
                driver.close();
                driver.switchTo().window(handles.iterator().next());
                Thread.sleep(200);
            }
        }
        driver.quit();
    }

    private static void writeExcel(List<ForexCalender> list, String year) {
        try {
            String fileName = "D:\\ForexFactory\\FxMEvents\\src\\FXM.xlsx";
            String sheetName = year.substring(1);
            FileInputStream inputStream = new FileInputStream(fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet(sheetName);
            int lastRowNum = sheet.getLastRowNum() + 1;
            int indx = 0;
            for (ForexCalender forexCalender : list) {
                Row row = sheet.createRow(lastRowNum++);
                if (!(forexCalender.eventDate.isEmpty())) {
                    row.createCell(0).setCellValue(forexCalender.eventDate);
                }
                if (!(forexCalender.time.isEmpty())) {
                    row.createCell(1).setCellValue(forexCalender.time);
                }
                if (!(forexCalender.currency.isEmpty())) {
                    row.createCell(2).setCellValue(forexCalender.currency);
                }
                if (!(forexCalender.eventName.isEmpty())) {
                    row.createCell(3).setCellValue(forexCalender.eventName);
                }
                if (!(forexCalender.actualNumber.isEmpty())) {
                    row.createCell(4).setCellValue(forexCalender.actualNumber);
                }
                if (!(forexCalender.forcastsNumber.isEmpty())) {
                    row.createCell(5).setCellValue(forexCalender.forcastsNumber);
                }
                if (!(forexCalender.previousNumbers.isEmpty())) {
                    row.createCell(6).setCellValue(forexCalender.previousNumbers);
                }
                indx++;
            }
            System.out.println((indx) + " - Row written");
            Thread.sleep(100);

            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
