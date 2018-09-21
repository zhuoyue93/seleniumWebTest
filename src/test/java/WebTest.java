import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class WebTest {
    @Test
    public  void main() throws InterruptedException {
        WebDriver driver;
//        System.setProperty("webdriver.chrome.driver", "D://web_selenium_test//chromedriver.exe");
//        driver = new ChromeDriver();
        System.setProperty("webdriver.ie.driver", "D://web_selenium_test//IEDriverServer.exe");
        driver = new InternetExplorerDriver();
        WebDriverWait wait = new WebDriverWait(driver, 1);
        new WebTest().picc_new(driver);
//        new WebTest().picc(driver);
    }
    private void picc(WebDriver driver) throws InterruptedException {
        driver.get("http://10.10.40.12:7014/piccallweb/haofeng3g.html");
        // 获取 网页的 title
        System.out.println("The testing page title is: " + driver.getTitle());
        //点击测试
        Select areaCode = new Select(driver.findElement(By.name("AreaCode")));
        areaCode.selectByVisibleText("广西");
        driver.findElement(By.name("button")).click();
        Thread.sleep(1000);


        //选中车险主页面
//        WebDriver frame = driver.switchTo().frame(1);
////        frame = frame.switchTo().frame(1);
//        driver=frame.switchTo().frame(0);
        driver.switchTo().defaultContent();

        //等待页面加载完成
        WebDriverWait wait = new WebDriverWait(driver, 5);
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("CarBrandModelName")));
        //录入“车辆品牌型号”
        driver.findElement(By.id("YEJ_LicenseNo")).sendKeys("qq");
        //点击“查询车型”
        driver.findElement(By.name("QueryCarBtn")).click();
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
