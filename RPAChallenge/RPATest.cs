using ExcelDataReader;
using OpenQA.Selenium;
using System.Collections;
using System.Data;
namespace RPAChallenge
{
    public class Tests
    {
        //variable for interacting with browser
        private IWebDriver driver;

        //variables containing web elements to search on the page
        //input fields
        private readonly By _firstName= By.XPath("//input[@ng-reflect-name='labelFirstName']");
        private readonly By _lastName = By.XPath("//input[@ng-reflect-name='labelLastName']");
        private readonly By _companyName = By.XPath("//input[@ng-reflect-name='labelCompanyName']");
        private readonly By _phoneNumber = By.XPath("//input[@ng-reflect-name='labelPhone']");
        private readonly By _address = By.XPath("//input[@ng-reflect-name='labelAddress']");
        private readonly By _email = By.XPath("//input[@ng-reflect-name='labelEmail']");
        private readonly By _roleInCompany = By.XPath("//input[@ng-reflect-name='labelRole']");

        //buttons
        private readonly By _startButton = By.XPath("//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']");
        private readonly By _submitButton = By.XPath("//input[@class='btn uiColorButton']");

        private static ArrayList ExcelFileReader()
        {
            //change encoding
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ArrayList excelData = new();
            //opening excel file to read data
            FileStream stream = File.Open("challenge.xlsx", FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            //reading from file and writing to ArrayList
            foreach (DataTable table in tables)
            {
                int k = 0;
                foreach (DataRow r in table.Rows)
                {
                    //don't write column names
                    if (k == 0)
                    {
                        k++;
                        continue;
                    }

                    for (var i = 0; i < 7; i++)
                    {
                        excelData.Add(r[i].ToString());
                    }
                }
            }
            //closing Excel file
            reader.Close();
            return excelData;
        }

        //opening the browser in full screen and going to the site
        [SetUp]
        public void Setup()
        {
            driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://rpachallenge.com/");
        }

        [Test]
        public void Test1()
        {
            //button to start rounds
            var start = driver.FindElement(_startButton);
            start.Click();

            //get ArrayList from Excel read function
            ArrayList data = ExcelFileReader();

            //additional variable for moving to the next line (next employee)
            int k = 0;

            //start rounds of filling in information
            for (var i = 0; i < 10; i++)
            {
                //Thread.Sleep(1000);
                var firstName = driver.FindElement(_firstName);
                firstName.SendKeys((string)data[0 + k]);

                //Thread.Sleep(1000);
                var lastName = driver.FindElement(_lastName);
                lastName.SendKeys((string)data[1 + k]);

                //Thread.Sleep(1000);
                var companyName = driver.FindElement(_companyName);
                companyName.SendKeys((string)data[2 + k]);

                //Thread.Sleep(1000);
                var phoneNumber = driver.FindElement(_phoneNumber);
                phoneNumber.SendKeys((string)data[6 + k]);

                //Thread.Sleep(1000);
                var address = driver.FindElement(_address);
                address.SendKeys((string)data[4 + k]);

                //Thread.Sleep(1000);
                var email = driver.FindElement(_email);
                email.SendKeys((string)data[5 + k]);

                //Thread.Sleep(1000);
                var roleInCompany = driver.FindElement(_roleInCompany);
                roleInCompany.SendKeys((string)data[3 + k]);

                //jump to a new line (to the data for the next round)
                k += 7;

                //button to go to the next round
                //Thread.Sleep(1000);
                var submit = driver.FindElement(_submitButton);
                submit.Click();
            }
        }

        //closing the browser with a delay to view the result
        [TearDown]
        public void TearDown()
        {
            Thread.Sleep(8000);
            driver.Quit();
        }
    }
}