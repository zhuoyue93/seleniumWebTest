import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.List;

public class KuaJingTest {
    @Test
    public  void test() throws InterruptedException {
        System.setProperty("webdriver.chrome.driver", "D://web_selenium_test//chromedriver.exe");
        //登录数据列表页面
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.gotenchina.com/amazon.html");

        //获取数据列表地址信息
        ArrayList<String[]> strings = new ArrayList<>();
        List<WebElement> elements = driver.findElements(By.xpath("//div[@class='goods5n']//a[@target='_blank']"));
        elements.forEach(element ->strings.add(new String[]{element.getAttribute("href"),""}));

        //获取产品英文名称
        strings.forEach(strings1 -> {
            driver.get(strings1[0]);
            WebElement element = driver.findElement(By.xpath("//div[@class='goods-tit']//h4"));
            strings1[1] =element.getText();
        });

        strings.forEach(strings2 -> System.out.println("{\""+strings2[0]+"\",\""+strings2[1]+"\"}"));
//        System.out.println(pageSource);
    }
}
