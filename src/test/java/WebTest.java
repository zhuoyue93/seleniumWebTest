import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import java.io.*;
import java.util.List;

public class WebTest {
    @Test
    public  void test() throws InterruptedException {
        WebDriver driver;
        //谷歌浏览器
        System.setProperty("webdriver.chrome.driver", "D://web_selenium_test//chromedriver.exe");
//        driver = new ChromeDriver();
        //IE浏览器
//        System.setProperty("webdriver.ie.driver", "D://web_selenium_test//IEDriverServer.exe");
//        driver = new InternetExplorerDriver();
//        WebDriverWait wait = new WebDriverWait(driver, 1);
//        String excelPath = "C:\\Users\\Administrator\\Desktop\\基础档案部1909单点方案11.xlsx";
        String excelPath = "C:\\Users\\Administrator\\Desktop\\基础档案部1909流量方案 - 副本.xlsx";

        readExcle(excelPath);
//        new WebTest().ncc(driver);
    }

    private void readExcle(String excelPath) {
        File excel = new File(excelPath);
        Workbook wb;
        try {
            FileInputStream in = new FileInputStream(excel);
            wb = new XSSFWorkbook(in);
            Sheet sheet = wb.getSheetAt(1);     //读取sheet 0

            int firstRowIndex = sheet.getFirstRowNum()+1;   //第一行是列名，所以不读
            int lastRowIndex = sheet.getLastRowNum();
//            System.out.println("firstRowIndex: "+firstRowIndex);
//            System.out.println("lastRowIndex: "+lastRowIndex);

            for(int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) {   //遍历行
//                System.out.println("rIndex: " + rIndex);
                Row row = sheet.getRow(rIndex);
                if (row != null) {
                    int firstCellIndex = row.getFirstCellNum();
                    int lastCellIndex = row.getLastCellNum();
                    for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {   //遍历列
                        Cell cell = row.getCell(cIndex);
                        if (cell != null) {
//                            System.out.println(cell.toString());

                        }
                    }
                    if(row.getCell(4).toString().contains("卓越")&row.getCell(3).toString().contains("打开节点")){
                        String s5 = row.getCell(2).toString();
                        String string =s5;
                        if(s5.contains(")")){
                            string=string.substring(3);
                        }

                        WebDriver driver=null;
                        OutputStream out=null;
                        try {
                            driver= new ChromeDriver();
                            String[] ncc = new WebTest().ncc(driver, string);
                            String[] s = ncc[0].split(" ");
                            boolean check=false;
                            if(s.length==41){
                                check=true;
                            }
                            String s1 = s[s.length - 18];//远程调用次数
                            Cell cell = row.getCell(9);//远程调用次数
                            if(cell.toString().trim().length()==0){
                                cell.setCellValue(s1);
                            }
                            if(Long.valueOf(s1)>10){
                                check=true;
                            }
                            String s2 = s[s.length - 7];//上行流量
                            Cell cell2 = row.getCell(8);//上行流量
                            if(cell2.toString().trim().length()==0){
                                cell2.setCellValue(s2);
                            }
                            if(Long.valueOf(s2)>150000){
                                check=true;
                            }
                            String s3 = s[s.length - 6];//下行流量
                            Cell cell3 = row.getCell(7);//下行流量
                            if(cell3.toString().trim().length()==0){
                                cell3.setCellValue(s3);
                            }
                            if(Long.valueOf(s3)>200000){
                                check=true;
                            }
                            String s4 = s[s.length - 12];//sql数量
                            Cell cell4 = row.getCell(10);//sql数量
                            if(cell4.toString().trim().length()==0){
                                cell4.setCellValue(s4);
                            }
                            if(Long.valueOf(s4)>200){
                                check=true;
                            }
                            if(check){
                                Cell cell5 = row.getCell(15);//备注
                                if(cell5.toString().trim().length()==0){
                                    cell5.setCellValue(ncc[1]);
                                }
                            }
                            // 创建文件输出流，准备输出电子表格：这个必须有，否则你在sheet上做的任何操作都不会有效
                            out =  new FileOutputStream(excelPath);
                            wb.write(out);
                        }catch (Exception e){
                            e.printStackTrace();
                        }finally {
                            driver.quit();
                            if(out != null){
                                out.flush();
                                out.close();
                            }
                        }
                    }
                }
            }


        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String[] ncc(WebDriver driver, String string) throws InterruptedException {
        String text=null;
        String currentUrl=null;
        driver.get("http://20.10.130.141:9520");
//        DesiredCapabilities ieCapabilities = DesiredCapabilities.internetExplorer();
//        ieCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
        // 获取 网页的 title
        System.out.println("The testing page title is: " + driver.getTitle());
        WebDriverWait wait = new WebDriverWait(driver, 5);
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("username")));
        WebElement username = driver.findElement(By.id("username"));
        username.clear();
        String usernamecode="y03";
        for (char a : usernamecode.toCharArray()) {
            username.sendKeys(String.valueOf(a));
        }
        WebElement password = driver.findElement(By.id("password"));
        password.clear();
        String passwordcode="123qwe";
        for (char a : passwordcode.toCharArray()) {
            password.sendKeys(String.valueOf(a));
        }
        WebElement loginBtn = driver.findElement(By.id("loginBtn"));
        loginBtn.click();
        wait.until(ExpectedConditions.presenceOfElementLocated(By.className("alert-ok")));
        driver.findElement(By.className("alert-ok")).click();

        wait.until(ExpectedConditions.presenceOfElementLocated(By.className("nc-workbench-icon-close")));
        driver.findElement(By.className("nc-workbench-icon-close")).click();
        wait.until(ExpectedConditions.presenceOfElementLocated(By.className("result-header-name")));
            ((ChromeDriver)driver).getMouse().click(null);
            wait.until(ExpectedConditions.not(ExpectedConditions.elementToBeClickable(By.className("result-header-name"))));
        WebElement until = wait.until(ExpectedConditions.elementToBeClickable(By.className("spr-record")));
        until.click();

//        driver.findElement(By.className("spr-record")).click();
        driver.findElement(By.className("nc-workbench-icon-close")).click();
        wait.until(ExpectedConditions.presenceOfElementLocated(By.className("result-header-name")));
        List<WebElement> elements = driver.findElements(By.xpath("//*/a"));
        for (WebElement element : elements) {
            if (string.equals(element.getText())){
                long start=System.currentTimeMillis();
                element.click();
                wait.until(ExpectedConditions.numberOfwindowsToBe(2));
                driver.switchTo().window(driver.getWindowHandles().toArray()[1].toString());
                wait.until(ExpectedConditions.titleContains(string));
                driver.findElement(By.className("spr-record")).click();
                wait.until(ExpectedConditions.numberOfwindowsToBe(3));
                driver.switchTo().window(driver.getWindowHandles().toArray()[2].toString());
                wait.until(ExpectedConditions.titleContains("SPR"));
                currentUrl = driver.getCurrentUrl();
                text = driver.findElements(By.xpath("//*/table")).get(2).getText();
                long end = System.currentTimeMillis();
                System.out.println(string+":操作耗时："+(end-start));
                break;
            }
        }
        return new String[]{text,currentUrl};
    }

    private void picc(WebDriver driver) throws InterruptedException {
        driver.get("http://10.10.40.13:7014/piccallweb/haofeng3g.html");
//        DesiredCapabilities ieCapabilities = DesiredCapabilities.internetExplorer();
//        ieCapabilities.setCapability(InternetExplorerDriver.INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS,true);
        // 获取 网页的 title
        System.out.println("The testing page title is: " + driver.getTitle());
        //点击测试
        Select areaCode = new Select(driver.findElement(By.name("AreaCode")));
        areaCode.selectByVisibleText("广西");
        driver.findElement(By.name("button")).click();
        Thread.sleep(1000);


        //选中车险主页面
        driver.switchTo().frame(1);
        driver.switchTo().frame(0);
        driver.switchTo().frame(0);

        //等待页面加载完成
        WebDriverWait wait = new WebDriverWait(driver, 5);
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CarBrandModelName")));
        //录入“车辆品牌型号”
        WebElement carBrandModelName = driver.findElement(By.id("CarBrandModelName"));
        carBrandModelName.clear();
        carBrandModelName.sendKeys("qq");
        //点击“查询车型”
        driver.findElement(By.name("QueryCarBtn")).click();
        //选择车型
        Select carkind = new Select(driver.findElement(By.name("CarQuerySelectList")));
        carkind.selectByIndex(2);
        //录入车牌
        WebElement yej_licenseNo = driver.findElement(By.id("YEJ_LicenseNo"));
        yej_licenseNo.sendKeys("201805");
        //录入初等日期
        WebElement enrollDate = driver.findElement(By.id("EnrollDate"));
        enrollDate.sendKeys("201805");
        driver.quit();
    }

    public void picc_new(WebDriver driver){
        driver.get("http://10.10.40.11:8101/tmgsp/html/haofeng3g.html");
        // 获取 网页的 title
        System.out.println("The testing page title is: " + driver.getTitle());
        //点击测试
        driver.findElement(By.name("button")).click();
        //选中车险主页面
        driver.switchTo().frame(0);
        //录入“车辆品牌型号”
        driver.findElement(By.xpath("//div/input[@id='carbrandmodelname']")).sendKeys("qq");
        //点击“查询车型”
        driver.findElement(By.id("select_car")).click();
        driver.quit();
    }

}
