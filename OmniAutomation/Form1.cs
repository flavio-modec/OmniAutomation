using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace OmniAutomation
{
    public partial class Form1 : Form
    {
        Options myOptions = new Options();
        public bool abort = false;
        static string myDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        System.Timers.Timer daily = new System.Timers.Timer(); //create timer

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Load Options
            myOptions.start(myDirectory);
            repPath.Text = myOptions.reportPath;
            xmlPath.Text = myOptions.xml004Path;
            unitCode.Text = myOptions.unitCode;
            genTime.Text = myOptions.genTime;

            //Configure timer
            daily.SynchronizingObject = this;
            daily.Elapsed += new System.Timers.ElapsedEventHandler(daily_Generate);
            bool t = daily.AutoReset;
            setTimer();
        }

        private void setTimer()
        {
            try
            {
                int genHour = Convert.ToInt16(myOptions.genTime.Substring(0, 2));
                int genMin = Convert.ToInt16(myOptions.genTime.Substring(3, 2));
                daily.Stop();
                if (DateTime.Now.Hour < genHour || (DateTime.Now.Hour == genHour && DateTime.Now.Minute < genMin))
                { daily.Interval = (DateTime.Today.AddHours(genHour).AddMinutes(genMin) - DateTime.Now).TotalMilliseconds; }
                else
                { daily.Interval = (DateTime.Today.AddDays(1).AddHours(genHour).AddMinutes(genMin) - DateTime.Now).TotalMilliseconds; }
                daily.Start();
            }
            catch { MessageBox.Show("Invalid Time!"); }
        }

        void daily_Generate(object sender, System.Timers.ElapsedEventArgs e)
        {
            daily.Stop();
            bool state = daily.Enabled;
            daily.Interval = 24 * 3600 * 1000; // 1 day in milliseconds
            daily.Start();
            getAll(); //get reports and configs
        }

        private void getReport_Click(object sender, EventArgs e)
        {
            getAll(); 
        }

        private void getAll()
        {
            string newFolder = myOptions.reportPath + "/Daily Report " + DateTime.Today.Date.ToString("dd-MMM-yyyy");
            GetReports GR = new GetReports(myDirectory, newFolder);
            XMLcreator xmlCreator = new XMLcreator();
            Config27 configGas;
            Config24 configOleo;
            String[] omniSN = new string[myOptions.nOmnis];

            if (System.IO.Directory.Exists(newFolder))
                System.IO.Directory.Delete(newFolder, true) ; //If the folder already exists, replace

            System.IO.DirectoryInfo todayDir = System.IO.Directory.CreateDirectory(newFolder);

            this.TopMost = true;

            for (ushort FC = 1; FC <= myOptions.nOmnis; FC++) //teste
            {
                status.Text = "FC" + Convert.ToString(FC) + " - Alarms...";
                this.Update();
                GR.GetAlarms(progressBar1, FC, 500);
                        
                status.Text = "FC" + Convert.ToString(FC) + " - Events...";
                this.Update();
                GR.GetEvents(progressBar1, FC, 150);
                        
                status.Text = "FC" + Convert.ToString(FC) + " - Daily...";
                this.Update();
                GR.GetDaily(progressBar1, FC);
                        
                if (FC <= myOptions.nOilOmnis)
                {
                    configOleo = new Config24(FC);
                    status.Text = "FC" + Convert.ToString(FC) + " - Downloading...";
                    this.Update();
                    configOleo.Download(progressBar1);
                    status.Text = "FC" + Convert.ToString(FC) + " - Saving...";
                    this.Update();
                    configOleo.Save(progressBar1, todayDir.FullName);
                    omniSN[FC-1] = configOleo.SN;
                }
                else
                {
                    configGas = new Config27(FC);
                    status.Text = "FC" + Convert.ToString(FC) + " - Downloading...";
                    this.Update();
                    configGas.Download(progressBar1);
                    status.Text = "FC" + Convert.ToString(FC) + " - Saving...";
                    this.Update();
                    configGas.Save(progressBar1, todayDir.FullName);
                    omniSN[FC-1] = configGas.SN;
                }
                        
            }

            status.Text = "Log...";
            LogCreator loger = new LogCreator();
            loger.createLog(myOptions, myDirectory);

            status.Text = "XML004...";
            
             xmlCreator.getXML004(myOptions.fiscalOmnis, omniSN, myOptions.unitCode, myOptions.xml004Path, todayDir.FullName); 
            //catch { MessageBox.Show("Error creating XML004! \n Check connection with flow computers and try again."); }
            status.Text = "Done.";

            this.TopMost = false;
        }

        private void apply_Click(object sender, EventArgs e)
        {
            myOptions.reportPath = repPath.Text;
            myOptions.xml004Path = xmlPath.Text;
            myOptions.unitCode = unitCode.Text;
            myOptions.genTime = genTime.Text;
            myOptions.save(myDirectory);
            setTimer();
            MessageBox.Show("Options were saved.");
        }


    }
}
