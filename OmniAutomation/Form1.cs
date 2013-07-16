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
        public Form1()
        {
            InitializeComponent();
        }

        System.Timers.Timer daily = new System.Timers.Timer(); //create timer

        private void Form1_Load(object sender, EventArgs e)
        {
            daily.SynchronizingObject = this;
            daily.Elapsed += new System.Timers.ElapsedEventHandler(daily_Generate);
            bool t = daily.AutoReset;
            setTimer(); 
        }

        private void applyTime_Click(object sender, EventArgs e)
        {
            setTimer();
        }

        private void setTimer()
        {
            try
            {
                int genHour = Convert.ToInt16(genTime.Text.Substring(0, 2));
                int genMin = Convert.ToInt16(genTime.Text.Substring(3, 2));
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
            const ushort nOmnis = 18; //number of flow computers
            const ushort nOilOmnis = 10; //number of Oil flow computers
            ushort[] FiscalOmnis = { 3, 4, 5, 11, 12, 13, 14, 15 }; //Selection of Fiscal flow computers
            GetReports GR = new GetReports();
            XMLcreator xmlCreator = new XMLcreator();
            Config27 configGas;
            Config24 configOleo;
            String[] omniSN = new string[nOmnis];

            string newFolder = path.Text + "/Daily Report " + DateTime.Today.Date.ToString("dd-MMM-yyyy");

            if (System.IO.Directory.Exists(newFolder))
                System.IO.Directory.Delete(newFolder, true) ; //If the folder already exists, replace
            System.IO.DirectoryInfo todayDir = System.IO.Directory.CreateDirectory(newFolder);

            string myDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            this.TopMost = true;

            for (ushort FC = 1; FC <= nOmnis; FC++) //teste
            {
                for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting data from Omni if it fails
                {
                    try
                    {
                        status.Text = "FC" + Convert.ToString(FC) + " - Alarms...";
                        this.Update();
                        GR.GetAlarms(progressBar1, FC, 500);
                        System.IO.File.Move(myDirectory + "/FC" + Convert.ToString(FC) + "_alarms.txt", todayDir.FullName + "/FC" + Convert.ToString(FC) + "_alarms.txt");

                        status.Text = "FC" + Convert.ToString(FC) + " - Events...";
                        this.Update();
                        GR.GetEvents(progressBar1, FC, 150);
                        System.IO.File.Move(myDirectory + "/FC" + Convert.ToString(FC) + "_events.txt", todayDir.FullName + "/FC" + Convert.ToString(FC) + "_events.txt");

                        status.Text = "FC" + Convert.ToString(FC) + " - Daily...";
                        this.Update();
                        GR.GetDaily(progressBar1, FC);
                        System.IO.File.Move(myDirectory + "/FC" + Convert.ToString(FC) + "_daily.txt", todayDir.FullName + "/FC" + Convert.ToString(FC) + "_daily.txt");

                        if (FC <= nOilOmnis)
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
                        break;
                    }
                    catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
                }
            }

            status.Text = "XML004...";
            
             xmlCreator.getXML004(FiscalOmnis, omniSN, unitCode.Text, xmlPath.Text, todayDir.FullName); 
            //catch { MessageBox.Show("Error creating XML004! \n Check connection with flow computers and try again."); }
            status.Text = "Done.";

            this.TopMost = false;
        }
    }
}
