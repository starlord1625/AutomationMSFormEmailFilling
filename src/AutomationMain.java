import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

//Download and add the following jar files in the lib folder:
//https://bit.ly/2SG4r3Y
//https://bit.ly/2Y6HRY9
//https://bit.ly/2LJ1leO
//https://bit.ly/2LHjeL9
//https://bit.ly/2ybqxBR

//Download and the jars in zip file
//https://jar-download.com/artifacts/org.seleniumhq.selenium/selenium-java/3.141.59/source-code

public class AutomationMain
{
    public static void main(String[] args) throws InterruptedException {
        //location of chrome driver
        //92 / 93 version from here : https://chromedriver.chromium.org/downloads
        System.setProperty("webdriver.chrome.driver","E:\\Softwares\\chromedriver_win32_92\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        //url of forms
        String baseUrl = "https://forms.office.com/Pages/DesignPage.aspx?auth_pvr=OrgId&auth_upn=harshalimi%40iitkgp.ac.in&origin=OfficeDotCom&lang=en-GB&route=Start#FormId=IrXbcQRXN0WfJWrS3NQnjYhOqajV03JEnLwM7xkC7h9UOTBXNTVXTDVES1M5VDZVRFIyMzBYVUZVVS4u&FlexPane=SendForm";

        // maximise the window
        driver.manage().window().maximize();

        //Deleting all the cookies
        driver.manage().deleteAllCookies();

        //Specifiying pageLoadTimeout and Implicit wait
        driver.manage().timeouts().pageLoadTimeout(40, TimeUnit.SECONDS);
        driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);

        //xpaths
        String shareButton="//button[@aria-label='Share']";
        String passwordFiller="//input[@placeholder='Password']";
        String yesButton="//input[@value='Yes']";
        String emailFiller="//input[@placeholder='Enter a name, group, or email address']";

        //password : to be filled
        String password="";

        //launch Chrome and direct it to the Base URL
        driver.get(baseUrl);

        Thread.sleep(10000);

        driver.findElement(By.xpath(passwordFiller)).sendKeys(password+ Keys.ENTER);

        Thread.sleep(5000);

        driver.findElement(By.xpath(yesButton)).click();

        Thread.sleep(5000);

        //click on share button
        //driver.findElement(By.xpath(shareButton)).click();
        try
        {
            File file = new File("E:\\info\\iitkgpDomains.xlsx");   //creating a new file instance and location of xlsx folder
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file
            itr.next();
            while (itr.hasNext())
            {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    //adding each email into filler and hitting enter
                    driver.findElement(By.xpath(emailFiller)).sendKeys(cell.getStringCellValue());
                    Thread.sleep(3000);
                    driver.findElement(By.xpath(emailFiller)).sendKeys(Keys.ENTER);
                }
            }
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}
