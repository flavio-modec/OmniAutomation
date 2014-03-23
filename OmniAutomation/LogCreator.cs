using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.IO;

namespace OmniAutomation
{
    class LogCreator
    {
        public void createLog(Options myOptions, string myDir)
        {
            string todayDir = myOptions.reportPath + "/Daily Report " + DateTime.Today.Date.ToString("dd-MMM-yyyy");
            string yesterdayDir = myOptions.reportPath + "/Daily Report " + DateTime.Today.AddDays(-1).Date.ToString("dd-MMM-yyyy");

            string todayValue;
            string yesterdayValue;

            ushort[] endLine24 = {18, 40, 11, 10, 10, 9, 9, 65};
            ushort[] endLine27 = {20, 47, 39, 11, 10, 32, 65};
            ushort addTab;

            Workbook todayConfig;
            Workbook yesterdayConfig;

            string logFile = "log_" + DateTime.Today.Date.ToString("dd-MMM-yyyy") + ".txt";
            TextWriter log = new StreamWriter(logFile);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            for (ushort FC = 1; FC <= myOptions.nOmnis; FC++)
            {
                
                //Compare Oil Config
                if (FC <= myOptions.nOilOmnis)
                {
                    try { todayConfig = excelApp.Workbooks.Open(todayDir + "/FC" + Convert.ToString(FC) + "_" + DateTime.Today.Date.ToString("dd-MMM-yyyy") + ".24.xls"); }
                    catch
                    {
                        log.WriteLine("FC" + FC.ToString() + "- Today configuration not found!");
                        continue;
                    }
                    try { yesterdayConfig = excelApp.Workbooks.Open(yesterdayDir + "/FC" + Convert.ToString(FC) + "_" + DateTime.Today.AddDays(-1).Date.ToString("dd-MMM-yyyy") + ".24.xls"); }
                    catch
                    {
                        todayConfig.Close(false);
                        log.WriteLine("FC" + FC.ToString() + "- Yesterday configuration not found!");
                        continue;
                    }
                    
                    //sheet config

                    for (ushort line = 2; line <= endLine24[0]; line++)
                    {
                        if (line == 7 || line == 8) //dont compare date and time
                            continue;
                        todayValue = Convert.ToString(todayConfig.Sheets[1].Cells[3][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[1].Cells[3][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + todayConfig.Sheets[1].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                    }
                    //sheets meter, tempertature, pressure and density
                    for (ushort meter = 1; meter <= 4; meter++)
                    {
                        for (ushort tab = 2; tab <= 5; tab++)
                        {
                            for (ushort line = 3; line <= endLine24[tab]; line++)
                            {
                                todayValue = Convert.ToString(todayConfig.Sheets[tab].Cells[meter + 2][line].Value);
                                yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[tab].Cells[meter + 2][line].Value);
                                if (!string.Equals(todayValue, yesterdayValue))
                                    log.WriteLine("FC" + FC.ToString() + " - Stream " + meter.ToString() + " - " + todayConfig.Sheets[tab].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                            }
                        }
                    }
                    //sheet products
                    for (ushort product = 1; product <= 12; product++)
                    {
                        for (ushort line = 3; line <= endLine24[5]; line++)
                        {
                            todayValue = Convert.ToString(todayConfig.Sheets[6].Cells[product + 2][line].Value);
                            yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[6].Cells[product + 2][line].Value);
                            if (!string.Equals(todayValue, yesterdayValue))
                                log.WriteLine("FC" + FC.ToString() + " - Product " + product.ToString() + " - " + todayConfig.Sheets[6].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                        }
                    }
                    //sheet prover
                    for (ushort line = 2; line <= endLine24[6]; line++)
                    {
                        todayValue = Convert.ToString(todayConfig.Sheets[7].Cells[3][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[7].Cells[3][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + todayConfig.Sheets[7].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                    }
                    //sheet statements
                    for (ushort line = 2; line <= endLine24[7]; line++)
                    {
                        todayValue = Convert.ToString(todayConfig.Sheets[8].Cells[3][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[8].Cells[3][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + "statement " + todayConfig.Sheets[8].Cells[2][line].Value + ": " + yesterdayValue + " to " + todayValue);
                    }
                    for (ushort line = 2; line <= endLine24[7]; line++)
                    {
                        todayValue = Convert.ToString(todayConfig.Sheets[8].Cells[6][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[8].Cells[6][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + "statement " + todayConfig.Sheets[8].Cells[5][line].Value + ": " + yesterdayValue + " to " + todayValue);
                    }
                }
                //Compare Gas Config
                else
                {
                    try { todayConfig = excelApp.Workbooks.Open(todayDir + "/FC" + Convert.ToString(FC) + "_" + DateTime.Today.Date.ToString("dd-MMM-yyyy") + ".27.xls"); }
                    catch
                    {
                        log.WriteLine("FC" + FC.ToString() + "- Today configuration not found!");
                        continue;
                    }
                    try { yesterdayConfig = excelApp.Workbooks.Open(yesterdayDir + "/FC" + Convert.ToString(FC) + "_" + DateTime.Today.AddDays(-1).Date.ToString("dd-MMM-yyyy") + ".27.xls"); }
                    catch
                    {
                        todayConfig.Close(false);
                        log.WriteLine("FC" + FC.ToString() + "- Yesterday configuration not found!");
                        continue;
                    }

                    //sheet config
                    for (ushort line = 2; line <= endLine27[0]; line++)
                    {
                        if (line == 7 || line == 8) //dont compare date and time
                            continue;
                        todayValue = Convert.ToString(todayConfig.Sheets[1].Cells[3][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[1].Cells[3][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + todayConfig.Sheets[1].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                    }
                    //sheets meter, tempertature, pressure
                    for (ushort meter = 1; meter <= 4; meter++)
                    {
                        for (ushort tab = 2; tab <= 5; tab++)
                        {
                            for (ushort line = 3; line <= endLine27[tab]; line++)
                            {
                                todayValue = Convert.ToString(todayConfig.Sheets[tab].Cells[meter + 2][line].Value);
                                yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[tab].Cells[meter + 2][line].Value);
                                if (!string.Equals(todayValue, yesterdayValue))
                                    log.WriteLine("FC" + FC.ToString() + " - Stream " + meter.ToString() + " - " + todayConfig.Sheets[tab].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                            }
                        }
                    }
                    //sheet products
                    for (ushort product = 1; product <= 4; product++)
                    {
                        for (ushort line = 3; line <= endLine27[5]; line++)
                        {
                            todayValue = Convert.ToString(todayConfig.Sheets[6].Cells[product + 2][line].Value);
                            yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[6].Cells[product + 2][line].Value);
                            if (!string.Equals(todayValue, yesterdayValue))
                                log.WriteLine("FC" + FC.ToString() + " - Product " + product.ToString() + " - " + todayConfig.Sheets[6].Cells[2][line].Value + yesterdayValue + " to " + todayValue);
                        }
                    }
                    //sheet statements
                    for (ushort line = 2; line <= endLine27[6]; line++)
                    {
                        todayValue = Convert.ToString(todayConfig.Sheets[7].Cells[3][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[7].Cells[3][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + "statement " + todayConfig.Sheets[7].Cells[2][line].Value + ": " + yesterdayValue + " to " + todayValue);
                    }
                    for (ushort line = 2; line <= endLine27[6]; line++)
                    {
                        todayValue = Convert.ToString(todayConfig.Sheets[7].Cells[6][line].Value);
                        yesterdayValue = Convert.ToString(yesterdayConfig.Sheets[7].Cells[6][line].Value);
                        if (!string.Equals(todayValue, yesterdayValue))
                            log.WriteLine("FC" + FC.ToString() + " - " + "statement " + todayConfig.Sheets[7].Cells[5][line].Value + ": " + yesterdayValue + " to " + todayValue);
                    }
                }
                todayConfig.Close(false);
                yesterdayConfig.Close(false);
            }
            excelApp.Quit();
            log.Close();
            System.IO.File.Move(myDir + "/" + logFile, todayDir + "/" + logFile);
            
        }
    }
}
