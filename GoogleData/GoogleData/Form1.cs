using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;



namespace GoogleData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;

        }
       


        private void button2_Click(object sender, EventArgs e)
        {
            int adet = listBox1.Items.Count;
            if (listBox1.Items.Count > 0)
            {
                Random rastgele = new Random();
                int sayi = rastgele.Next();

                string dosyayolu;

                dosyayolu = "D:\\" + sayi.ToString() + ".xlsx";
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = true;

                object Missing = Type.Missing;

                Workbook calismakitabi = excelapp.Workbooks.Add(Missing);
                
                Worksheet sheet1 = (Worksheet)calismakitabi.Sheets[1];
                int sutun = 1;
                int satir = 1;
                for (int i = 0; i < adet; i++)
                {
                    for (int j = 0; j < 1; j++)
                    {
                        Range myrange1 = (Range)sheet1.Cells[satir + i, sutun + j];
                        myrange1.Value2 = listBox1.Items[i] == null ? "" : listBox1.Items[i];
                        myrange1.Select();

                    }
                }

               
                
                listBox1.Items.Clear();
            }
        }

        private void yazAra()
        {
            
            
            try
            {
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                ChromeOptions options = new ChromeOptions();
                options.AddExcludedArgument("enable-automation");
                options.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(service,options);
                driver.Navigate().GoToUrl("https://www.google.com/");
                Thread.Sleep(5000);
                IWebElement element = driver.FindElement(By.XPath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input"));
                element.SendKeys(textBox1.Text);
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[2]/div[2]/center/input[1]")).Click();
                Thread.Sleep(2000);

                for (int i = 1; i < 2; i++)
                {
                    try
                    {
                        string span = driver.FindElement(By.ClassName("LrzXr")).Text;
                        string ad = driver.FindElement(By.ClassName("qrShPb")).Text;
                        string tel = driver.FindElement(By.ClassName("zdqRlf")).Text;
                        string hour = driver.FindElement(By.ClassName("JjSWRd")).Text;


                        listBox1.Items.Add("Firma Adı: " + ad + "  Adresi: " + span + " Telefon:" + tel + "Çalışma Durumu: "+hour );

                    }
                    catch (Exception)
                    {

                        string span = driver.FindElement(By.ClassName("LrzXr")).Text;
                        string ad = driver.FindElement(By.ClassName("qrShPb")).Text;
                        string tel = driver.FindElement(By.ClassName("zdqRlf")).Text;

                        listBox1.Items.Add("Firma Adı: " + ad + "  Adresi: " + span + " Telefon:" + tel );
                    }

                   
                    








                }
                driver.Close();




            }
            catch (Exception)
            {
                MessageBox.Show("Çalışma Saati Belli Değil!");
                


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Thread th = new Thread(yazAra);
            th.Start();

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void Form1_Load(object sender, EventArgs e)
        {
        }
    }
}
