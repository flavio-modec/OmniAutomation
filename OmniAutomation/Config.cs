using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using Microsoft.Office.Interop.Excel;
using Modbus.Device;

namespace OmniAutomation
{
    class Config24
    {
        private ushort _FC;
        private string _IP;
        private string _SN;
        private general24 _general24;
        private flowmeterPulse[] _flowmeterPulse = new flowmeterPulse[4];
        private temperature[] _temperature = new temperature[4];
        private pressure[] _pressure = new pressure[4];
        private density[] _density = new density[4];
        private product24[] _product24 = new product24[12];
        private prover _prover;
        private statments _statments;

        public Config24(ushort FC)
        {
            _FC = FC;
            _IP = "10.10.1." + Convert.ToString(FC + 10);
        }

        public string SN { get { return _SN; } }

        public void Download(System.Windows.Forms.ProgressBar progress)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting config from Omni if it fails
            {
                try
                {
                    Download_attempt(progress);
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        public void Save(System.Windows.Forms.ProgressBar progress, string directory)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting config from Omni if it fails
            {
                try
                {
                    Save_attempt(progress, directory);
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        private void Download_attempt(System.Windows.Forms.ProgressBar progress)
        {
            float[] floatRange;

            //Connect
            TcpClient client = new TcpClient(_IP, 502);
            client.ReceiveTimeout = 1000;
            ModbusIpMaster master = ModbusIpMaster.CreateIp(client);
            OmniIpMaster omniMaster = new OmniIpMaster(master);

            _SN = omniMaster.ReadString(14099, 1);

            _general24.ID = omniMaster.ReadString(4836, 1);
            _general24.location = omniMaster.ReadString(4842, 5);
            _general24.company = omniMaster.ReadString(4837, 5);
            _general24.dateFormat = omniMaster.ReadInt16(3842, 1)[0];
            _general24.sysDate = omniMaster.ReadString(4847, 1);
            _general24.sysTime = omniMaster.ReadString(4848, 1);
            _general24.volUnit = omniMaster.ReadInt16(3097, 1)[0];
            _general24.pressUnit = omniMaster.ReadInt16(13071, 1)[0];
            _general24.atmPressure = omniMaster.ReadFloat(7891, 1)[0];
            _general24.roll = omniMaster.ReadInt16(3098, 1)[0];
            _general24.volDecPlaces = omniMaster.ReadInt16(13386, 1)[0];
            _general24.massDecPlaces = omniMaster.ReadInt16(13388, 1)[0];
            //progressBar
            progress.Value = 8;

            for (ushort i = 0; i < 4; i++)
            {
                _flowmeterPulse[i].IOpoint = omniMaster.ReadInt16((ushort)(13001 + 13 * i), 1)[0];
                _flowmeterPulse[i].ID = omniMaster.ReadString((ushort)(4114 + 100 * i), 1);
                _flowmeterPulse[i].model = omniMaster.ReadString((ushort)(4113 + 100 * i), 1);
                _flowmeterPulse[i].size = omniMaster.ReadString((ushort)(4112 + 100 * i), 1);
                _flowmeterPulse[i].SN = omniMaster.ReadString((ushort)(4111 + 100 * i), 1);
                _flowmeterPulse[i].grossFS = omniMaster.ReadFloat((ushort)(17176 + 4 * i), 1)[0];
                _flowmeterPulse[i].massFS = omniMaster.ReadFloat((ushort)(17177 + 4 * i), 1)[0];
                _flowmeterPulse[i].lowLimit = omniMaster.ReadFloat((ushort)(7161 + 100 * i), 1)[0];
                _flowmeterPulse[i].highLimit = omniMaster.ReadFloat((ushort)(7162 + 100 * i), 1)[0];
                _flowmeterPulse[i].activeFreq = omniMaster.ReadInt16((ushort)(3106 + 100 * i), 1)[0];
                _flowmeterPulse[i].alarmInactive = omniMaster.ReadInt16((ushort)(13437 + i), 1)[0];
                _flowmeterPulse[i].pulseFidelity = omniMaster.ReadInt16((ushort)(13413 + 13 * i), 1)[0];
                _flowmeterPulse[i].kFactor = omniMaster.ReadFloat((ushort)(17501 + 100 * i), 12);
                _flowmeterPulse[i].freq = omniMaster.ReadInt16((ushort)(3122 + 100 * i), 12);

                floatRange = omniMaster.ReadFloat((ushort)(7163 + 100 * i), 15);

                _temperature[i].IOpoint = omniMaster.ReadInt16((ushort)(13002 + 13 * i), 1)[0];
                _temperature[i].tag = omniMaster.ReadString((ushort)(4117 + 100 * i), 1);
                _temperature[i].type = omniMaster.ReadString((ushort)(13003 + 13 * i), 1);
                _temperature[i].lowAlarm = floatRange[0];
                _temperature[i].highAlarm = floatRange[1];
                _temperature[i].overrideCode = omniMaster.ReadInt16((ushort)(3101 + 100 * i), 1)[0];
                _temperature[i].overrideValue = floatRange[2];
                _temperature[i].zeroScale = floatRange[3];
                _temperature[i].fullScale = floatRange[4];

                _pressure[i].IOpoint = omniMaster.ReadInt16((ushort)(13004 + 13 * i), 1)[0];
                _pressure[i].tag = omniMaster.ReadString((ushort)(4118 + 100 * i), 1);
                _pressure[i].lowAlarm = floatRange[5];
                _pressure[i].highAlarm = floatRange[6];
                _pressure[i].overrideCode = omniMaster.ReadInt16((ushort)(3102 + 100 * i), 1)[0];
                _pressure[i].overrideValue = floatRange[7];
                _pressure[i].zeroScale = floatRange[8];
                _pressure[i].fullScale = floatRange[9];

                _density[i].IOpoint = omniMaster.ReadInt16((ushort)(13005 + 13 * i), 1)[0];
                _density[i].tag = omniMaster.ReadString((ushort)(4119 + 100 * i), 1);
                _density[i].lowAlarm = floatRange[10];
                _density[i].highAlarm = floatRange[11];
                _density[i].overrideCode = omniMaster.ReadInt16((ushort)(3102 + 100 * i), 1)[0];
                _density[i].overrideValue = floatRange[12];
                _density[i].zeroScale = floatRange[13];
                _density[i].fullScale = floatRange[14];
                progress.Increment(4);
            }
            
            for (ushort j = 0; j < 12; j++)
            {
                _product24[j].name = omniMaster.ReadString((ushort)(4820 + j), 1);
                _product24[j].MF = new Int32[4];
                for (ushort i = 0; i < 4; i++)
                {
                    _product24[j].MF[i] = omniMaster.ReadInt32((ushort)(5121 + 100 * i + j), 1)[0];
                }
                _product24[j].algorithm = omniMaster.ReadInt16((ushort)(3813 + j), 1)[0];
                _product24[j].refDensityOverride = omniMaster.ReadFloat((ushort)(7822 + j), 1)[0];
                _product24[j].refTemperature = omniMaster.ReadFloat((ushort)(17219 + j), 1)[0];
                progress.Increment(1);
            }
            //progressBar
            progress.Value = 32;

            _prover.type = omniMaster.ReadInt16(3921, 1)[0];
            _prover.volume = omniMaster.ReadFloat(7919, 1)[0];
            _prover.runsToAverage = omniMaster.ReadInt16(3913, 1)[0];
            _prover.maxRuns = omniMaster.ReadInt16(3914, 1)[0];
            _prover.basedOn = omniMaster.ReadInt16(3928, 1)[0];
            _prover.repeatability = omniMaster.ReadFloat(7927, 1)[0];
            _prover.autoImplementMF = omniMaster.ReadInt16(3922, 1)[0];
            _prover.maxDeviation = omniMaster.ReadFloat(7928, 1)[0];
            //progressBar
            progress.Value = 40;

            //statements
            _statments.boolStatment = new string[64];
            _statments.boolRemark = new string[64];
            _statments.varStatment = new string[64];
            _statments.varRemark = new string[64];
            for (ushort i = 0; i < 48; i++)
            {
                _statments.boolStatment[i] = omniMaster.ReadString((ushort)(14001 + i), 1);
                _statments.boolRemark[i] = omniMaster.ReadString((ushort)(14101 + i), 1);
                _statments.varStatment[i] = omniMaster.ReadString((ushort)(14051 + i), 1);
                _statments.varRemark[i] = omniMaster.ReadString((ushort)(14151 + i), 1);
            }
            for (ushort i = 0; i < 16; i++)
            {
                _statments.boolStatment[i + 48] = omniMaster.ReadString((ushort)(14201 + i), 1);
                _statments.boolRemark[i + 48] = omniMaster.ReadString((ushort)(14241 + i), 1);
                _statments.varStatment[i + 48] = omniMaster.ReadString((ushort)(14221 + i), 1);
                _statments.varRemark[i + 48] = omniMaster.ReadString((ushort)(14261 + i), 1);
            }
            //progressBar
            progress.Value = 63;//complete
        }

        public void Save_attempt(System.Windows.Forms.ProgressBar progress, string directory)
        {
            Application excelApp = new Application();
            string myDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            Workbook Omni24 = excelApp.Workbooks.Open(myDirectory + "/Omni.24.xls");

            Worksheet generalTab = Omni24.Sheets[1];
            generalTab.Cells[3][2] = _general24.ID;
            generalTab.Cells[3][3] = _general24.location;
            generalTab.Cells[3][4] = _general24.company;
            generalTab.Cells[3][6] = _general24.dateFormat;
            generalTab.Cells[3][7] = _general24.sysDate;
            generalTab.Cells[3][8] = _general24.sysTime;
            generalTab.Cells[3][11] = _general24.volUnit;
            generalTab.Cells[3][12] = _general24.pressUnit;
            generalTab.Cells[3][13] = _general24.atmPressure;
            generalTab.Cells[3][16] = _general24.roll;
            generalTab.Cells[3][17] = _general24.volDecPlaces;
            generalTab.Cells[3][18] = _general24.massDecPlaces;
            progress.Value = 8;

            Worksheet flowmeterTab = Omni24.Sheets[2];
            for (ushort i = 0; i < 4; i++)
            {
                if (_flowmeterPulse[i].IOpoint == 0) { break; }
                flowmeterTab.Cells[3 + i][3] = _flowmeterPulse[i].IOpoint;
                flowmeterTab.Cells[3 + i][4] = _flowmeterPulse[i].ID;
                flowmeterTab.Cells[3 + i][5] = _flowmeterPulse[i].model;
                flowmeterTab.Cells[3 + i][6] = _flowmeterPulse[i].size;
                flowmeterTab.Cells[3 + i][7] = _flowmeterPulse[i].SN;
                flowmeterTab.Cells[3 + i][8] = _flowmeterPulse[i].grossFS;
                flowmeterTab.Cells[3 + i][9] = _flowmeterPulse[i].massFS;
                flowmeterTab.Cells[3 + i][10] = _flowmeterPulse[i].lowLimit;
                flowmeterTab.Cells[3 + i][11] = _flowmeterPulse[i].highLimit;
                flowmeterTab.Cells[3 + i][12] = _flowmeterPulse[i].activeFreq;
                flowmeterTab.Cells[3 + i][13] = _flowmeterPulse[i].pulseFidelity;
                for (ushort k = 0; k < 12; k++)
                {
                    flowmeterTab.Cells[3 + i][17 + k] = _flowmeterPulse[i].kFactor[k];
                    flowmeterTab.Cells[3 + i][29 + k] = _flowmeterPulse[i].freq[k];
                }
            }
            progress.Value = 16;

            Worksheet temperatureTab = Omni24.Sheets[3];
            for (ushort i = 0; i < 4; i++)
            {
                temperatureTab.Cells[3 + i][3] = _temperature[i].IOpoint;
                temperatureTab.Cells[3 + i][4] = _temperature[i].tag;
                temperatureTab.Cells[3 + i][5] = _temperature[i].type;
                temperatureTab.Cells[3 + i][6] = _temperature[i].lowAlarm;
                temperatureTab.Cells[3 + i][7] = _temperature[i].highAlarm;
                temperatureTab.Cells[3 + i][8] = _temperature[i].overrideCode;
                temperatureTab.Cells[3 + i][9] = _temperature[i].overrideValue;
                temperatureTab.Cells[3 + i][10] = _temperature[i].zeroScale;
                temperatureTab.Cells[3 + i][11] = _temperature[i].fullScale;
            }
            progress.Value = 20;

            Worksheet pressureTab = Omni24.Sheets[4];
            for (ushort i = 0; i < 4; i++)
            {
                pressureTab.Cells[3 + i][3] = _pressure[i].IOpoint;
                pressureTab.Cells[3 + i][4] = _pressure[i].tag;
                pressureTab.Cells[3 + i][5] = _pressure[i].lowAlarm;
                pressureTab.Cells[3 + i][6] = _pressure[i].highAlarm;
                pressureTab.Cells[3 + i][7] = _pressure[i].overrideCode;
                pressureTab.Cells[3 + i][8] = _pressure[i].overrideValue;
                pressureTab.Cells[3 + i][9] = _pressure[i].zeroScale;
                pressureTab.Cells[3 + i][10] = _pressure[i].fullScale;
            }
            progress.Value = 24;

            Worksheet densityTab = Omni24.Sheets[5];
            for (ushort i = 0; i < 4; i++)
            {
                densityTab.Cells[3 + i][3] = _density[i].IOpoint;
                densityTab.Cells[3 + i][4] = _density[i].tag;
                densityTab.Cells[3 + i][5] = _density[i].lowAlarm;
                densityTab.Cells[3 + i][6] = _density[i].highAlarm;
                densityTab.Cells[3 + i][7] = _density[i].overrideCode;
                densityTab.Cells[3 + i][8] = _density[i].overrideValue;
                densityTab.Cells[3 + i][9] = _density[i].zeroScale;
                densityTab.Cells[3 + i][10] = _density[i].fullScale;
            }
            progress.Value = 28;
            
            Worksheet productsTab = Omni24.Sheets[6];
            for (ushort j = 0; j < 12; j++)
            {
                productsTab.Cells[3 + j][2] = _product24[j].name;
                for (ushort i = 0; i < 4; i++)
                {
                    productsTab.Cells[3 + j][3 + i] = _product24[j].MF[i];
                }
                productsTab.Cells[3 + j][7] = _product24[j].algorithm;
                productsTab.Cells[3 + j][8] = _product24[j].refDensityOverride;
                productsTab.Cells[3 + j][9] = _product24[j].refTemperature;
            }
            progress.Value = 36;

            Worksheet proverTab = Omni24.Sheets[7];
            proverTab.Cells[3][2] = _prover.type;
            proverTab.Cells[3][3] = _prover.volume;
            proverTab.Cells[3][4] = _prover.runsToAverage;
            proverTab.Cells[3][5] = _prover.maxRuns;
            proverTab.Cells[3][6] = _prover.basedOn;
            proverTab.Cells[3][7] = _prover.repeatability;
            proverTab.Cells[3][8] = _prover.autoImplementMF;
            proverTab.Cells[3][9] = _prover.maxDeviation;
            progress.Value = 40;

            Worksheet statementsTab = Omni24.Sheets[8];
            for (ushort i = 0; i < 64; i++)
            {
                statementsTab.Cells[3][2 + i] = _statments.boolStatment[i];
                statementsTab.Cells[4][2 + i] = _statments.boolRemark[i];
                statementsTab.Cells[6][2 + i] = _statments.varStatment[i];
                statementsTab.Cells[7][2 + i] = _statments.varRemark[i];
            }
            progress.Value = 60;

            Omni24.SaveCopyAs(directory + "/FC" + Convert.ToString(_FC) + ".24.xls");
            //Omni24.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, directory + "/FC" + Convert.ToString(_FC) + ".24.pdf");
            Omni24.Close(false);
            excelApp.Quit();
            progress.Value = 63;//Complete
        }
    }

    class Config27
    {
        private ushort _FC;
        private string _IP;
        private string _SN;
        private general27 _general27;
        private ushort[] _flowmenterType = new ushort[4];
        private flowmeterPulse[] _flowmeterPulse = new flowmeterPulse[4];
        private flowmeterDP[] _flowmeterDP = new flowmeterDP[4];
        private temperature[] _temperature = new temperature[4];
        private pressure[] _pressure = new pressure[4];
        private product27[] _product27 = new product27[4];
        private statments _statments;

        public Config27(ushort FC)
        {
            _FC = FC;
            _IP = "10.10.1." + Convert.ToString(FC + 10);
        }

        public string ID { get { return _general27.ID; } }
        public string SN { get { return _SN; } }

        public void Download(System.Windows.Forms.ProgressBar progress)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting config from Omni if it fails
            {
                try
                {
                    Download_attempt(progress);
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        public void Save(System.Windows.Forms.ProgressBar progress, string directory)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting config from Omni if it fails
            {
                try
                {
                    Save_attempt(progress, directory);
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        private void Download_attempt(System.Windows.Forms.ProgressBar progress)
        {
            float[] floatRange;
            //progress bar
            //Form1 popForm = new Form1();
            //popForm.Show();
            //Connect

            TcpClient client = new TcpClient(_IP, 502);
            client.ReceiveTimeout = 1000;
            ModbusIpMaster master = ModbusIpMaster.CreateIp(client);
            OmniIpMaster omniMaster = new OmniIpMaster(master);

            _SN = omniMaster.ReadString(14099, 1);

            _general27.ID = omniMaster.ReadString(4836, 1);
            _general27.location = omniMaster.ReadString(4842, 5);
            _general27.company = omniMaster.ReadString(4837, 5);
            _general27.dateFormat = omniMaster.ReadInt16(3842, 1)[0];
            _general27.sysDate = omniMaster.ReadString(4847, 1);
            _general27.sysTime = omniMaster.ReadString(4848, 1);
            _general27.DPUnit = omniMaster.ReadInt16(13072, 1)[0];
            _general27.pressUnit = omniMaster.ReadInt16(13071, 1)[0];
            _general27.atmPressure = omniMaster.ReadFloat(7891, 1)[0];
            _general27.roll = omniMaster.ReadInt16(3098, 1)[0];
            _general27.grossDecPlaces = omniMaster.ReadInt16(13386, 1)[0];
            _general27.netDecPlaces = omniMaster.ReadInt16(13387, 1)[0];
            _general27.massDecPlaces = omniMaster.ReadInt16(13388, 1)[0];
            _general27.energyDecPlaces = omniMaster.ReadInt16(13389, 1)[0];
            //progress bar
            progress.Value = 4;

            for (ushort i = 0; i < 4; i++)
            {
                _flowmenterType[i] = omniMaster.ReadInt16((ushort)(3108 + 100 * i), 1)[0];
                if (_flowmenterType[i] == 1) //if type = pulse
                {
                    _flowmeterPulse[i].IOpoint = omniMaster.ReadInt16((ushort)(13001 + 13 * i), 1)[0];
                    _flowmeterPulse[i].ID = omniMaster.ReadString((ushort)(4114 + 100 * i), 1);
                    _flowmeterPulse[i].model = omniMaster.ReadString((ushort)(4113 + 100 * i), 1);
                    _flowmeterPulse[i].size = omniMaster.ReadString((ushort)(4112 + 100 * i), 1);
                    _flowmeterPulse[i].SN = omniMaster.ReadString((ushort)(4111 + 100 * i), 1);
                    _flowmeterPulse[i].grossFS = omniMaster.ReadFloat((ushort)(17176 + 4 * i), 1)[0];
                    _flowmeterPulse[i].massFS = omniMaster.ReadFloat((ushort)(17177 + 4 * i), 1)[0];
                    _flowmeterPulse[i].lowLimit = omniMaster.ReadFloat((ushort)(7161 + 100 * i), 1)[0];
                    _flowmeterPulse[i].highLimit = omniMaster.ReadFloat((ushort)(7162 + 100 * i), 1)[0];
                    _flowmeterPulse[i].activeFreq = omniMaster.ReadInt16((ushort)(3106 + 100 * i), 1)[0];
                    _flowmeterPulse[i].alarmInactive = omniMaster.ReadInt16((ushort)(13437 + i), 1)[0];
                    _flowmeterPulse[i].pulseFidelity = omniMaster.ReadInt16((ushort)(13413 + 13 * i), 1)[0];
                    _flowmeterPulse[i].kFactor = omniMaster.ReadFloat((ushort)(17501 + 100 * i), 12);
                    _flowmeterPulse[i].freq = omniMaster.ReadInt16((ushort)(3122 + 100 * i), 12);
                }
                if (_flowmenterType[i] == 0 || _flowmenterType[i] == 2) //if type = DP
                {
                    _flowmeterDP[i].DPlowIOpoint = omniMaster.ReadInt16((ushort)(13011 + 13 * i), 1)[0];
                    _flowmeterDP[i].DPmidIOpoint = omniMaster.ReadInt16((ushort)(13059 + i), 1)[0];
                    _flowmeterDP[i].DPhighIOpoint = omniMaster.ReadInt16((ushort)(13012 + 13 * i), 1)[0];
                    _flowmeterDP[i].ID = omniMaster.ReadString((ushort)(4114 + 100 * i), 1);
                    _flowmeterDP[i].model = omniMaster.ReadString((ushort)(4113 + 100 * i), 1);
                    _flowmeterDP[i].size = omniMaster.ReadString((ushort)(4112 + 100 * i), 1);
                    _flowmeterDP[i].SN = omniMaster.ReadString((ushort)(4111 + 100 * i), 1);
                    _flowmeterDP[i].grossFS = omniMaster.ReadFloat((ushort)(17176 + 4 * i), 1)[0];
                    _flowmeterDP[i].massFS = omniMaster.ReadFloat((ushort)(17177 + 4 * i), 1)[0];
                    _flowmeterDP[i].lowLimit = omniMaster.ReadFloat((ushort)(7161 + 100 * i), 1)[0];
                    _flowmeterDP[i].highLimit = omniMaster.ReadFloat((ushort)(7162 + 100 * i), 1)[0];

                    _flowmeterDP[i].dpLowTag = omniMaster.ReadString((ushort)(4115 + 100 * i), 1);
                    _flowmeterDP[i].dpLowZeroScale = omniMaster.ReadFloat((ushort)(7155 + 100 * i), 1)[0];
                    _flowmeterDP[i].dpLowFullScale = omniMaster.ReadFloat((ushort)(7156 + 100 * i), 1)[0];

                    _flowmeterDP[i].dpMidTag = omniMaster.ReadString((ushort)(4128 + 100 * i), 1);
                    _flowmeterDP[i].dpMidZeroScale = omniMaster.ReadFloat((ushort)(17400 + 2 * i), 1)[0];
                    _flowmeterDP[i].dpMidFullScale = omniMaster.ReadFloat((ushort)(17401 + 2 * i), 1)[0];

                    _flowmeterDP[i].dpHighTag = omniMaster.ReadString((ushort)(4116 + 100 * i), 1);
                    _flowmeterDP[i].dpHighZeroScale = omniMaster.ReadFloat((ushort)(7157 + 100 * i), 1)[0];
                    _flowmeterDP[i].dpHighFullScale = omniMaster.ReadFloat((ushort)(7158 + 100 * i), 1)[0];

                    _flowmeterDP[i].overrideCode = omniMaster.ReadInt16((ushort)(3109 + 100 * i), 1)[0];
                    _flowmeterDP[i].overrideValue = omniMaster.ReadFloat((ushort)(7154 + 100 * i), 1)[0];
                    _flowmeterDP[i].lowAlarm = omniMaster.ReadFloat((ushort)(7152 + 100 * i), 1)[0];
                    _flowmeterDP[i].highAlarm = omniMaster.ReadFloat((ushort)(7153 + 100 * i), 1)[0];
                    _flowmeterDP[i].switchHigh = omniMaster.ReadFloat((ushort)(7159 + 100 * i), 1)[0];
                    _flowmeterDP[i].switchLow = omniMaster.ReadFloat((ushort)(7160 + 100 * i), 1)[0];

                    _flowmeterDP[i].orificeDiam = omniMaster.ReadFloat((ushort)(7145 + 100 * i), 1)[0];
                    _flowmeterDP[i].orificeExpCoef = omniMaster.ReadFloat((ushort)(7146 + 100 * i), 1)[0];
                    _flowmeterDP[i].orificeRefTemp = omniMaster.ReadFloat((ushort)(7147 + 100 * i), 1)[0];

                    _flowmeterDP[i].pipeDiam = omniMaster.ReadFloat((ushort)(7148 + 100 * i), 1)[0];
                    _flowmeterDP[i].pipeExpCoef = omniMaster.ReadFloat((ushort)(7149 + 100 * i), 1)[0];
                    _flowmeterDP[i].pipeRefTemp = omniMaster.ReadFloat((ushort)(7150 + 100 * i), 1)[0];

                    _flowmeterDP[i].deviceType = omniMaster.ReadInt16((ushort)(3112 + 100 * i), 1)[0];
                }

                floatRange = omniMaster.ReadFloat((ushort)(7163 + 100 * i), 15);

                _temperature[i].IOpoint = omniMaster.ReadInt16((ushort)(13002 + 13 * i), 1)[0];
                _temperature[i].tag = omniMaster.ReadString((ushort)(4117 + 100 * i), 1);
                _temperature[i].type = omniMaster.ReadString((ushort)(13003 + 13 * i), 1);
                _temperature[i].lowAlarm = floatRange[0];
                _temperature[i].highAlarm = floatRange[1];
                _temperature[i].overrideCode = omniMaster.ReadInt16((ushort)(3101 + 100 * i), 1)[0];
                _temperature[i].overrideValue = floatRange[2];
                _temperature[i].zeroScale = floatRange[3];
                _temperature[i].fullScale = floatRange[4];

                _pressure[i].IOpoint = omniMaster.ReadInt16((ushort)(13004 + 13 * i), 1)[0];
                _pressure[i].tag = omniMaster.ReadString((ushort)(4118 + 100 * i), 1);
                _pressure[i].lowAlarm = floatRange[5];
                _pressure[i].highAlarm = floatRange[6];
                _pressure[i].overrideCode = omniMaster.ReadInt16((ushort)(3102 + 100 * i), 1)[0];
                _pressure[i].overrideValue = floatRange[7];
                _pressure[i].zeroScale = floatRange[8];
                _pressure[i].fullScale = floatRange[9];
            }
            //progress bar
            progress.Value = 24;

            for (ushort j = 0; j < 4; j++)
            {
                _product27[j].type = omniMaster.ReadInt16((ushort)(3813 + j), 1)[0];
                _product27[j].name = omniMaster.ReadString((ushort)(4820 + j), 1);
                _product27[j].viscosity = omniMaster.ReadFloat((ushort)(17251 + 30 * j), 1)[0];
                _product27[j].isoentropicExp = omniMaster.ReadFloat((ushort)(17252 + 30 * j), 1)[0];
                _product27[j].heatingValue = omniMaster.ReadFloat((ushort)(17253 + 30 * j), 1)[0];
                _product27[j].relativeDens = omniMaster.ReadFloat((ushort)(17254 + 30 * j), 1)[0];
                _product27[j].densMethod = omniMaster.ReadInt16((ushort)(3817 + j), 1)[0];

                _product27[j].cromathography = new float[22];
                for (ushort element = 0; element < 21; element++)
                {
                    _product27[j].cromathography[element] = omniMaster.ReadFloat((ushort)(17230 + element + 30 * j), 1)[0];
                    
                }
                _product27[j].cromathography[21] = omniMaster.ReadFloat((ushort)(17258 + 30 * j), 1)[0]; //neo-pentane
                progress.Increment(4);
            }

            //statements

            _statments.boolStatment = new string[64];
            _statments.boolRemark = new string[64];
            _statments.varStatment = new string[64];
            _statments.varRemark = new string[64];
            for (ushort i = 0; i < 48; i++)
            {
                _statments.boolStatment[i] = omniMaster.ReadString((ushort)(14001 + i), 1);
                _statments.boolRemark[i] = omniMaster.ReadString((ushort)(14101 + i), 1);
                _statments.varStatment[i] = omniMaster.ReadString((ushort)(14051 + i), 1);
                _statments.varRemark[i] = omniMaster.ReadString((ushort)(14151 + i), 1);
                
            }
            for (ushort i=0; i<16;i++)
            {
                _statments.boolStatment[i+48] = omniMaster.ReadString((ushort)(14201 + i), 1);
                _statments.boolRemark[i+48] = omniMaster.ReadString((ushort)(14241 + i), 1);
                _statments.varStatment[i+48] = omniMaster.ReadString((ushort)(14221 + i), 1);
                _statments.varRemark[i+48] = omniMaster.ReadString((ushort)(14261 + i), 1);
                
            }
            //progress bar
            progress.Value = 63; //complete
            //popForm.Close();
        }
        private void Save_attempt(System.Windows.Forms.ProgressBar progress, string directory)
        {
            progress.Value = 0; 

            Application excelApp = new Application();
            string myDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            Workbook Omni27 = excelApp.Workbooks.Open(myDirectory + "/Omni.27.xls");

            Worksheet generalTab = Omni27.Sheets[1];
            generalTab.Cells[3][2] = _general27.ID;
            generalTab.Cells[3][3] = _general27.location;
            generalTab.Cells[3][4] = _general27.company;
            generalTab.Cells[3][6] = _general27.dateFormat;
            generalTab.Cells[3][7] = _general27.sysDate;
            generalTab.Cells[3][8] = _general27.sysTime;
            generalTab.Cells[3][11] = _general27.DPUnit;
            generalTab.Cells[3][12] = _general27.pressUnit;
            generalTab.Cells[3][13] = _general27.atmPressure;
            generalTab.Cells[3][16] = _general27.roll;
            generalTab.Cells[3][17] = _general27.grossDecPlaces;
            generalTab.Cells[3][18] = _general27.netDecPlaces;
            generalTab.Cells[3][19] = _general27.massDecPlaces;
            generalTab.Cells[3][20] = _general27.energyDecPlaces;
            //progress bar
            progress.Value = 8;

            //if (_flowmenterType. == 0) { Omni27.Sheets[3] ; }
            Worksheet flowmeterDPTab = Omni27.Sheets[2];
            Worksheet flowmeterPulseTab = Omni27.Sheets[3];
            bool existDP = false;
            bool existPulse = false;
            for (ushort i = 0; i < 4; i++)
            {
                if (_flowmenterType[i] == 1)
                {
                    if (_flowmeterPulse[i].IOpoint == 0) { break; }
                    existPulse = true;
                    flowmeterPulseTab.Cells[3 + i][3] = _flowmeterPulse[i].IOpoint;
                    flowmeterPulseTab.Cells[3 + i][4] = _flowmeterPulse[i].ID;
                    flowmeterPulseTab.Cells[3 + i][5] = _flowmeterPulse[i].model;
                    flowmeterPulseTab.Cells[3 + i][6] = _flowmeterPulse[i].size;
                    flowmeterPulseTab.Cells[3 + i][7] = _flowmeterPulse[i].SN;
                    flowmeterPulseTab.Cells[3 + i][8] = _flowmeterPulse[i].grossFS;
                    flowmeterPulseTab.Cells[3 + i][9] = _flowmeterPulse[i].massFS;
                    flowmeterPulseTab.Cells[3 + i][10] = _flowmeterPulse[i].lowLimit;
                    flowmeterPulseTab.Cells[3 + i][11] = _flowmeterPulse[i].highLimit;
                    flowmeterPulseTab.Cells[3 + i][12] = _flowmeterPulse[i].activeFreq;
                    flowmeterPulseTab.Cells[3 + i][13] = _flowmeterPulse[i].pulseFidelity;
                    if (_flowmeterPulse[i].kFactor != null)
                    {
                        for (ushort k = 0; k < 12; k++)
                        {
                            flowmeterPulseTab.Cells[3 + i][16 + k] = _flowmeterPulse[i].kFactor[k];
                            flowmeterPulseTab.Cells[3 + i][28 + k] = _flowmeterPulse[i].freq[k];
                        }
                    }
                }

                if (_flowmenterType[i] == 0 || _flowmenterType[i] == 2)
                {
                    if ((_flowmeterDP[i].DPlowIOpoint == 0) && (_flowmeterDP[i].DPhighIOpoint == 0)) {break;}
                    existDP = true;
                    flowmeterDPTab.Cells[3 + i][3] = _flowmeterDP[i].DPlowIOpoint;
                    flowmeterDPTab.Cells[3 + i][4] = _flowmeterDP[i].DPmidIOpoint;
                    flowmeterDPTab.Cells[3 + i][5] = _flowmeterDP[i].DPhighIOpoint;
                    flowmeterDPTab.Cells[3 + i][6] = _flowmeterDP[i].ID;
                    flowmeterDPTab.Cells[3 + i][7] = _flowmeterDP[i].model;
                    flowmeterDPTab.Cells[3 + i][8] = _flowmeterDP[i].size;
                    flowmeterDPTab.Cells[3 + i][9] = _flowmeterDP[i].SN;
                    flowmeterDPTab.Cells[3 + i][10] = _flowmeterDP[i].grossFS;
                    flowmeterDPTab.Cells[3 + i][11] = _flowmeterDP[i].massFS;
                    flowmeterDPTab.Cells[3 + i][12] = _flowmeterDP[i].lowLimit;
                    flowmeterDPTab.Cells[3 + i][13] = _flowmeterDP[i].highLimit;

                    flowmeterDPTab.Cells[3 + i][17] = _flowmeterDP[i].dpLowTag;
                    flowmeterDPTab.Cells[3 + i][18] = _flowmeterDP[i].dpLowZeroScale;
                    flowmeterDPTab.Cells[3 + i][19] = _flowmeterDP[i].dpLowFullScale;

                    flowmeterDPTab.Cells[3 + i][21] = _flowmeterDP[i].dpMidTag;
                    flowmeterDPTab.Cells[3 + i][22] = _flowmeterDP[i].dpMidZeroScale;
                    flowmeterDPTab.Cells[3 + i][23] = _flowmeterDP[i].dpMidFullScale;

                    flowmeterDPTab.Cells[3 + i][25] = _flowmeterDP[i].dpHighTag;
                    flowmeterDPTab.Cells[3 + i][26] = _flowmeterDP[i].dpHighZeroScale;
                    flowmeterDPTab.Cells[3 + i][27] = _flowmeterDP[i].dpHighFullScale;

                    flowmeterDPTab.Cells[3 + i][29] = _flowmeterDP[i].overrideCode;
                    flowmeterDPTab.Cells[3 + i][30] = _flowmeterDP[i].overrideValue;

                    flowmeterDPTab.Cells[3 + i][32] = _flowmeterDP[i].lowAlarm;
                    flowmeterDPTab.Cells[3 + i][33] = _flowmeterDP[i].highAlarm;

                    flowmeterDPTab.Cells[3 + i][35] = _flowmeterDP[i].switchHigh;
                    flowmeterDPTab.Cells[3 + i][36] = _flowmeterDP[i].switchLow;

                    flowmeterDPTab.Cells[3 + i][40] = _flowmeterDP[i].orificeDiam;
                    flowmeterDPTab.Cells[3 + i][41] = _flowmeterDP[i].orificeExpCoef;
                    flowmeterDPTab.Cells[3 + i][42] = _flowmeterDP[i].orificeRefTemp;

                    flowmeterDPTab.Cells[3 + i][44] = _flowmeterDP[i].pipeDiam;
                    flowmeterDPTab.Cells[3 + i][45] = _flowmeterDP[i].pipeExpCoef;
                    flowmeterDPTab.Cells[3 + i][46] = _flowmeterDP[i].pipeRefTemp;

                    flowmeterDPTab.Cells[3 + i][47] = _flowmeterDP[i].deviceType;
                }
            }
            if (!existDP) { flowmeterDPTab.Visible = XlSheetVisibility.xlSheetHidden; }
            if (!existPulse) { flowmeterPulseTab.Visible = XlSheetVisibility.xlSheetHidden; }
            //progress bar
            progress.Value = 16;

            Worksheet temperatureTab = Omni27.Sheets[4];
            for (ushort i = 0; i < 4; i++)
            {
                temperatureTab.Cells[3 + i][3] = _temperature[i].IOpoint;
                temperatureTab.Cells[3 + i][4] = _temperature[i].tag;
                temperatureTab.Cells[3 + i][5] = _temperature[i].type;
                temperatureTab.Cells[3 + i][6] = _temperature[i].lowAlarm;
                temperatureTab.Cells[3 + i][7] = _temperature[i].highAlarm;
                temperatureTab.Cells[3 + i][8] = _temperature[i].overrideCode;
                temperatureTab.Cells[3 + i][9] = _temperature[i].overrideValue;
                temperatureTab.Cells[3 + i][10] = _temperature[i].zeroScale;
                temperatureTab.Cells[3 + i][11] = _temperature[i].fullScale;
            }
            //progress bar
            progress.Value = 24;

            Worksheet pressureTab = Omni27.Sheets[5];
            for (ushort i = 0; i < 4; i++)
            {
                pressureTab.Cells[3 + i][3] = _pressure[i].IOpoint;
                pressureTab.Cells[3 + i][4] = _pressure[i].tag;
                pressureTab.Cells[3 + i][5] = _pressure[i].lowAlarm;
                pressureTab.Cells[3 + i][6] = _pressure[i].highAlarm;
                pressureTab.Cells[3 + i][7] = _pressure[i].overrideCode;
                pressureTab.Cells[3 + i][8] = _pressure[i].overrideValue;
                pressureTab.Cells[3 + i][9] = _pressure[i].zeroScale;
                pressureTab.Cells[3 + i][10] = _pressure[i].fullScale;
            }
            //progress bar
            progress.Value = 32;
            
            Worksheet productsTab = Omni27.Sheets[6];
            for (ushort j = 0; j < 4; j++)
            {
                productsTab.Cells[3 + j][2] = _product27[j].type;
                productsTab.Cells[3 + j][3] = _product27[j].name;
                productsTab.Cells[3 + j][4] = _product27[j].viscosity;
                productsTab.Cells[3 + j][5] = _product27[j].isoentropicExp;
                productsTab.Cells[3 + j][6] = _product27[j].heatingValue;
                productsTab.Cells[3 + j][7] = _product27[j].relativeDens;
                productsTab.Cells[3 + j][8] = _product27[j].densMethod;

                if (_product27[j].cromathography != null)
                {
                    for (ushort element = 0; element < 22; element++)
                    {
                        productsTab.Cells[3 + j][11 + element] = _product27[j].cromathography[element];
                    }
                }
            }
            //progress bar
            progress.Value = 44;

            //statements
            Worksheet statementsTab = Omni27.Sheets[7];
            for (ushort i = 0; i < 64; i++)
            {
                statementsTab.Cells[3][2 + i] = _statments.boolStatment[i];
                statementsTab.Cells[4][2 + i] = _statments.boolRemark[i];
                statementsTab.Cells[6][2 + i] = _statments.varStatment[i];
                statementsTab.Cells[7][2 + i] = _statments.varRemark[i];
            }
            
            progress.Value = 59;

            Omni27.SaveCopyAs(directory+"/FC"+Convert.ToString(_FC)+".27.xls");
            //Omni27.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, directory + "/FC" + Convert.ToString(_FC) + ".27.pdf");
            Omni27.Close(false);
            excelApp.Quit();
            
            progress.Value = 63; //Complete
        }
         
    }
    public struct general24
    {
        public string ID;
        public string location;
        public string company;
        public ushort dateFormat;
        public string sysDate;
        public string sysTime;
        public ushort volUnit;
        public ushort pressUnit;
        public float atmPressure;
        public ushort roll;
        public ushort volDecPlaces;
        public ushort massDecPlaces;
    }
    public struct general27
    {
        public string ID;
        public string location;
        public string company;
        public ushort dateFormat;
        public string sysDate;
        public string sysTime;
        public ushort DPUnit;
        public ushort pressUnit;
        public float atmPressure;
        public ushort roll;
        public ushort grossDecPlaces;
        public ushort netDecPlaces;
        public ushort massDecPlaces;
        public ushort energyDecPlaces;
    }
    public struct flowmeterPulse
    {
        public ushort IOpoint;
        public string ID;
        public string model;
        public string size;
        public string SN;
        public float grossFS;
        public float massFS;
        public float lowLimit;
        public float highLimit;
        public ushort activeFreq;
        public ushort alarmInactive;
        public ushort pulseFidelity;
        public float[] kFactor;
        public ushort[] freq;
    }
    public struct flowmeterDP
    {
            public ushort DPlowIOpoint;
            public ushort DPmidIOpoint;
            public ushort DPhighIOpoint;
            public string ID;
            public string model;
            public string size;
            public string SN;
            public float grossFS;
            public float massFS;
            public float lowLimit;
            public float highLimit;
            public string dpLowTag;
            public float dpLowZeroScale;
            public float dpLowFullScale;
            public string dpMidTag;
            public float dpMidZeroScale;
            public float dpMidFullScale;
            public string dpHighTag;
            public float dpHighZeroScale;
            public float dpHighFullScale;
            public ushort overrideCode;
            public float overrideValue;
            public float lowAlarm;
            public float highAlarm;
            public float switchHigh;
            public float switchLow;
            public float orificeDiam;
            public float orificeExpCoef;
            public float orificeRefTemp;
            public float pipeDiam;
            public float pipeExpCoef;
            public float pipeRefTemp;
            public ushort deviceType;
        }
    public struct temperature
    {
            public ushort IOpoint;
            public string tag;
            public string type;
            public float lowAlarm;
            public float highAlarm;
            public ushort overrideCode;
            public float overrideValue;
            public float zeroScale;
            public float fullScale;
    }
    public struct pressure
    {
            public ushort IOpoint;
            public string tag;
            public float lowAlarm;
            public float highAlarm;
            public ushort overrideCode;
            public float overrideValue;
            public float zeroScale;
            public float fullScale;
     }
    public struct density
    {
            public ushort IOpoint;
            public string tag;
            public float lowAlarm;
            public float highAlarm;
            public ushort overrideCode;
            public float overrideValue;
            public float zeroScale;
            public float fullScale;
    }
    public struct product24
    {
            public string name;
            public int[] MF;
            public ushort algorithm;
            public float refDensityOverride;
            public float refTemperature;
     }
    public struct product27
    {
            public ushort type;
            public string name;
            public float viscosity;
            public float isoentropicExp;
            public float heatingValue;
            public float relativeDens;
            public ushort densMethod;
            public float[] cromathography;
    }
    public struct prover
    {
            public ushort type;
            public float volume;
            public ushort runsToAverage;
            public ushort maxRuns;
            public ushort basedOn;
            public float repeatability;
            public ushort autoImplementMF;
            public float maxDeviation;
    }
    public struct statments
    {
            public string[] boolStatment;
            public string[] boolRemark;
            public string[] varStatment;
            public string[] varRemark;
    }
}