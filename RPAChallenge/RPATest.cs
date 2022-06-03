using ExcelDataReader;
using OpenQA.Selenium;
using System.Collections;
using System.Data;
namespace RPAChallenge
{
    public class Tests
    {
        //переменная для взаимодействия с браузером
        private IWebDriver driver;

        //переменные, содержащие веб-элементы для поиска на странице
        //поля ввода
        private readonly By _firstName= By.XPath("//input[@ng-reflect-name='labelFirstName']");
        private readonly By _lastName = By.XPath("//input[@ng-reflect-name='labelLastName']");
        private readonly By _companyName = By.XPath("//input[@ng-reflect-name='labelCompanyName']");
        private readonly By _phoneNumber = By.XPath("//input[@ng-reflect-name='labelPhone']");
        private readonly By _address = By.XPath("//input[@ng-reflect-name='labelAddress']");
        private readonly By _email = By.XPath("//input[@ng-reflect-name='labelEmail']");
        private readonly By _roleInCompany = By.XPath("//input[@ng-reflect-name='labelRole']");

        //кнопки
        private readonly By _startButton = By.XPath("//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']");
        private readonly By _submitButton = By.XPath("//input[@class='btn uiColorButton']");

        private static ArrayList ExcelFileReader()
        {
            //смена кодировки
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ArrayList excelData = new();
            //открытие Excel файла для чтения данных
            FileStream stream = File.Open("challenge.xlsx", FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            //чтение из файла и запись в ArrayList
            foreach (DataTable table in tables)
            {
                int k = 0;
                foreach (DataRow r in table.Rows)
                {
                    //Не записываем названия столбцов
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
            //закрытие файла 
            reader.Close();
            return excelData;
        }
           
        //Открытие браузера на полный экран и переход на сайт
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
            //Кнопка для начала раундов
            var start = driver.FindElement(_startButton);
            start.Click();

            //Возвращаем ArrayList из функции чтения Excel
            ArrayList data = ExcelFileReader();

            //Вспомогательная переменная для перехода к следующей строке(следующему сотруднику)
            int k = 0;
            //Начинаем раунды заполнения информации
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

                //Переход на новую строку(к данным для следующего раунда)
                k += 7;

                //Кнопка для перехода к следующему раунду
                //Thread.Sleep(1000);
                var submit = driver.FindElement(_submitButton);
                submit.Click();
            }
        }

        //Закрытие браузера с задержкой для просмотра результата
        [TearDown]
        public void TearDown()
        {
            Thread.Sleep(8000);
            driver.Quit();
        }
    }
}