using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using UgetsTest.ExcelModels;
using excel = Microsoft.Office.Interop.Excel;



namespace UgetsTest
{
    public class Tests
    {

        IWebDriver driver = null;
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(""); // We couldn't share webpage url, it is forbidden.
            Thread.Sleep(9999);
            
        }

        [Test]
        public void Test1()
        {
            string[] file_paths = Directory.GetFiles(@"C:\Users\batuh\source\repos\UgetsTest\UgetsTest\excelFiles\"); //It holds all excel files path in the ExcelFiles folder.
            
            int packetRow = 1;
           
            // test is made for all excel folders in the excelFiles folder respectively
            foreach ( var path in file_paths) 
            {
                
            // data in the excel files is transfered to the ExcelModel class
            List<ExcelModel> excelresults = ReadFromExcel(path);  
               
                // isFirstRow variable is a indicator that determines whether the program is in the first row
                bool isFirstRow = true;
                
                // all data is read and performed in the ExcelModel class
                foreach (ExcelModel model in excelresults)
            {
                // ">" is an indicator that distinct packets from the commands. If packet name starts with ">" it is a packet name so the program first search that packet and open its parameters.
                // packet names that don't start with ">" is parameters that belong to that packet name. Program insert these parameters to the web page.
                if (!model.tcMnemonic.StartsWith(">"))
                {

                        // it closes the previous packet table when the next packet has started. If the program is in the first row then there is no packet that has to be closed.
                        if (isFirstRow == false)
                        {
                            clickPacketParameterIcon(packetRow);
                        }
                    
                    // it writes and enters the packet name that starts with ">" on the search box then opens the packet parameters.
                    EnterSearchAndClick(model.tcMnemonic);
                    
                    // every packet has a time value on the packet row. It sets the time value.
                    SetTimeValue(model.Value, packetRow);
                    
                    Thread.Sleep(300);

                    // it locates the packet row's open/close icon and open the packet parameters    
                    IWebElement element = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td[5]"));  
                    Thread.Sleep(300);
                    element.Click();

                    packetRow++;


                }
                
                // This part of code executes the packet name doesn't start with ">". The data contains InputType that indicates type of the input. 
                else
                {
                    
                    if (model.InputType == "3")
                    {
              
                        SetDropdownValue(model.VariableName, model.Value, packetRow);

                    }
                    else if (model.InputType == "1")
                    {
                        SetTextValue(model.VariableName, model.Value, packetRow);
                    }
                    else if (model.InputType == "4")
                        {

                            clickPacketParameterIcon(packetRow);

                            DragDropElement(model);

                        }

                    isFirstRow = false;

                }
                    

                }

               


            }

            Thread.Sleep(500);
          
           
            Assert.Pass();
            driver.Close();

        }

        // It reads data from excelfiles and instert them to the excelmodel class than returns a list that holds all excelmodel objects.
        public List<ExcelModel> ReadFromExcel(string file)
        {
            
            List<ExcelModel> resultList = new List<ExcelModel>();



            excel.Application app = new excel.Application();
            
            // giving file path is opened
            excel.Workbook workbook = app.Workbooks.Open(Path.Combine(@"" + file));
            
            //data is in the first sheet everytime for every excelfiles
            excel.Worksheet worksheet = workbook.Sheets[1];



            //int excelRowRange = worksheet.UsedRange.Rows.Count;
            //int excelColumnRange = worksheet.UsedRange.Columns.Count;


            //The nested loop looks up every cell on the excelfile. The range can be determined with the previous comments, but we preferred to finish the loop when the next cell is null .
            for (int row = 2; row<=200 ; row++)
            {
                ExcelModel model = new ExcelModel();
                for (int column = 1; column < 5; column++)
                {
                    excel.Range range = (excel.Range)worksheet.Cells[row, column];
                    if (column == 1)
                    {
                        if (range.Value != null)
                        {
                            model.tcMnemonic = range.Value.ToString();
                        }
                        else
                        {
                            goto LoopEnd;
                        }
                    }
                    else if (column == 2)
                    {
                        
                            model.VariableName = range.Value.ToString();
                    }
                    else if (column == 3)
                    {
                        
                            model.Value = range.Value.ToString();
                    }
                    else if (column == 4)
                    {
                        
                            model.InputType = range.Value.ToString();
                    }
                    
                }
                resultList.Add(model);

            }
            LoopEnd:
                workbook.Close();
            app.Quit();
            return resultList;


        }

        // it closes and opens the packet parameters
        public void clickPacketParameterIcon(int counter)
        {
            
            IWebElement packetIcon = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + (counter-1) + "]/td[5]"));
            Thread.Sleep(300);
            packetIcon.Click();
            Thread.Sleep(300);
            
        }

        // it takes the packet name as a parameter then search the packet name and click on it.
        public void EnterSearchAndClick(string tcMnemonic)
        {

            // telemetry search gives us the locate of the searchbox
            IWebElement searchBox = driver.FindElement(By.Id("telemetrySearch"));

            // it paste the packet name to the searchbox
            searchBox.SendKeys(tcMnemonic);
            Thread.Sleep(500);

            // it click the packet name among the results 
            driver.FindElement(By.XPath("//div[@data-mnemonic='" + tcMnemonic + "']")).Click();
            Thread.Sleep(500);

            // lastly searchbox is cleared
            searchBox.Click();
            searchBox.Clear();
        }

        // it sets the time value on the packet row
        public void SetTimeValue(string input, int counter)

        {
            Thread.Sleep(500);

            // it locates the inputbox that is inserted time value
            IWebElement timeBox = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + counter + "]/td[2]/input"));
            
            Thread.Sleep(500);
            timeBox.Click();
            Thread.Sleep(500);
            timeBox.Clear();
            timeBox.SendKeys(input);
            Thread.Sleep(500);


        }

        // it clicked the given value in the dropdown list
        public void SetDropdownValue(string targetName, string value, int packetRow)
        {
            int dropdownRow = 2;

            
            IWebElement label = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[2]/label"));

            // it takes the first parameter name in the row
            string getname = label.Text;

            // it search the correct dropdown row from the name of the row
            while (targetName != getname)
            {
                dropdownRow++;
                label = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[" + dropdownRow + "]/label"));
                getname = label.Text;

            }
            
            // it clicks the dropdown box and opens the dropdown list
            
            IWebElement dropdown = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[" + dropdownRow + "]/div[1]/div"));
            dropdown.Click();
            Thread.Sleep(300);

            // it clicks the given value in the dropdown list
            dropdown = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[" + dropdownRow + "]/div[1]/div/div/table/tbody/tr[@data-value='" + value + "']"));
            dropdown.Click();

            Thread.Sleep(300);

        }

        // it inserts the text value to the textbox 
        public void SetTextValue(string targetName, string value, int packetRow)



        {
            int textRow = 2;

            
            IWebElement label = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[2]/label"));
            
            // it takes the first parameter name in the row
            string getname = label.Text;

            // it search the correct text row from the name of the row
            while (targetName != getname)
            {
                textRow++;
                label = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[" + textRow + "]/label"));
                getname = label.Text;

            }
            

            Thread.Sleep(300);
            // it clicks correct textbox on the parameterlist
            IWebElement textBox = driver.FindElement(By.XPath("//*[@id='tclTable']/tbody/tr[" + packetRow + "]/td/div/div[" + textRow + "]/div[1]/input"));
            textBox.Click();
            Thread.Sleep(300);
            // clear the texbox
            textBox.Clear();
            Thread.Sleep(300);
            //insert the value
            textBox.SendKeys(value);




        }
        // Webpage has drag and drop feature betweeen the packet rows. This function can be performed drag and drop operations between two packet row.
        public void DragDropElement(ExcelModel model)
        {
            string tCMnemonicStr = model.tcMnemonic;
            if (model.tcMnemonic.StartsWith(">"))
            {
                tCMnemonicStr = model.tcMnemonic.Remove(0, 1);
            }

            //WebElement on which drag and drop operation needs to be performed
            IWebElement fromElement = driver.FindElement(By.XPath("//table[@id='tclTable']//following::tr[@data-mnemonic='" + tCMnemonicStr + "'][1]//td[@title='Sırayı Değiştir'][1]"));

            //WebElement to which the above object is dropped
            IWebElement toElement = driver.FindElement(By.XPath("//table[@id='tclTable']//following::tr[@data-mnemonic='" + model.VariableName + "'][1]"));

            //Creating object of Actions class to build composite actions
            Actions builder = new Actions(driver);

            //Building a drag and drop action
            var dragAndDrop = builder.ClickAndHold(fromElement).Pause(TimeSpan.FromMilliseconds(1000)).MoveToElement(toElement, 20,-80).Pause(TimeSpan.FromMilliseconds(1000)).Release().Pause(TimeSpan.FromMilliseconds(1000)).Build();

            //Performing the drag and drop action
            dragAndDrop.Perform();

           
        }

    }


}