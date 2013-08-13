using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Sockets;
using System.Windows.Forms;
using Modbus;
using Modbus.Device;
using Modbus.Message;
using Modbus.Utility;
using Modbus.Data;
using Modbus.IntegrationTests.CustomMessages;

namespace OmniAutomation
{
    class GetReports
    {
        private string _myDir;
        private string _todayDir;

        public string myDir { get { return _myDir; } }
        public string todayDir { get { return _todayDir; } }

        public GetReports(string myDir, string todayDir)
        {
            _myDir = myDir;
            _todayDir = todayDir;
        }

        public void GetDaily(ProgressBar progress, ushort FC)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting daily from Omni if it fails
            {
                try
                {
                    GetDailyReport(progress, FC);
                    System.IO.File.Move(myDir + "/FC" + Convert.ToString(FC) + "_daily.txt", todayDir + "/FC" + Convert.ToString(FC) + "_daily.txt");
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        private void GetDailyReport(ProgressBar progress, ushort FC)
        {
            progress.Value = 0;
            //Connect
            string OmniIP = "10.10.1." + Convert.ToString(FC + 10);
            TcpClient client = new TcpClient(OmniIP, 502);
            client.ReceiveTimeout = 500;
            ModbusIpMaster master = ModbusIpMaster.CreateIp(client);
            
            //Create text file
            TextWriter txt = new StreamWriter("FC" + Convert.ToString(FC) + "_daily.txt");
            
            //Read buffer
            int endIndex;
            for (ushort pack = 0; pack < 20; pack++)
            {
                try  //send request to Omni for the current packet
                {
                    CustomReadBufferRequest reqBuffer = new CustomReadBufferRequest(65, 1, 9301, pack);
                    CustomReadBufferResponse packet = master.ExecuteCustomMessage<CustomReadBufferResponse>(reqBuffer);
                    //stop if find end of file
                    endIndex = packet.StrData.IndexOf(Convert.ToChar(26));//get index of end of file if exists
                    if (endIndex != -1) //if end of file was found
                    {
                        txt.Write(packet.StrData.Substring(0, endIndex));
                        break;
                    }
                    txt.Write(packet.StrData);
                    progress.Increment(1);
                }
                catch { break; }
            }
            progress.Value = 63;
            txt.Close();
            client.Close();
        }

        public void GetAlarms(ProgressBar progress, ushort FC, ushort quantity)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting daily from Omni if it fails
            {
                try
                {
                    GetAlarmsEvents(progress, FC, quantity, true);
                    System.IO.File.Move(myDir + "/FC" + Convert.ToString(FC) + "_alarms.txt", todayDir + "/FC" + Convert.ToString(FC) + "_alarms.txt");
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        public void GetEvents(ProgressBar progress, ushort FC, ushort quantity)
        {
            for (ushort attempt = 0; attempt < 3; attempt++) //try 3 times getting daily from Omni if it fails
            {
                try
                {
                    GetAlarmsEvents(progress, FC, quantity, false);
                    System.IO.File.Move(myDir + "/FC" + Convert.ToString(FC) + "_events.txt", todayDir + "/FC" + Convert.ToString(FC) + "_events.txt");
                    break;
                }
                catch { System.Threading.Thread.Sleep(500); } //wait 500ms before trying again
            }
        }

        private void GetAlarmsEvents(ProgressBar progress, ushort FC, ushort quantity, bool getAlarms)
        {
            progress.Value = 0;

            string OmniIP = "10.10.1." + Convert.ToString(FC + 10);
            TcpClient client = new TcpClient(OmniIP, 502);
            client.ReceiveTimeout = 500;
            ModbusIpMaster master = ModbusIpMaster.CreateIp(client);
            
            //Create text file
            string fileName;
            if (getAlarms) { fileName = "FC" + Convert.ToString(FC) + "_alarms.txt"; }
            else { fileName = "FC" + Convert.ToString(FC) + "_events.txt"; }
            TextWriter txt = new StreamWriter(fileName);

            //Number of events to load
            master.WriteSingleRegister((byte)1, (ushort)3769, quantity);

            //Send command to load Events on buffer
            byte commandByte;
            if (getAlarms) { commandByte = 0x10; }
            else { commandByte = 0x80; }
            //byte[] invCommandBytes = new byte[] { 0, commandNibble, 0, 0 };
            RegisterCollection invCommand = new RegisterCollection(new byte[] { 0, commandByte, 0, 0 });
            CustomWriteMultipleRegistersRequest reqCommand = new CustomWriteMultipleRegistersRequest(16, 1, 15129, invCommand);
            master.ExecuteCustomMessage<CustomWriteMultipleRegistersResponse>(reqCommand);

            //Wait buffer is ready
            CustomReadHoldingRegistersResponse cmdReg;
            do
            {
                CustomReadHoldingRegistersRequest readCmd = new CustomReadHoldingRegistersRequest(3, 1, 15129, 1);
                cmdReg = master.ExecuteCustomMessage<CustomReadHoldingRegistersResponse>(readCmd);
                Console.WriteLine(Convert.ToString(cmdReg.Data[1]));
            } while (cmdReg.Data[1] != 0);

            //Read buffer
            int endIndex;
            int CRIndex;
            DateTime timeStamp;
            IFormatProvider dateFormat = new System.Globalization.CultureInfo("en-GB");
            for (ushort i = 0; i < 600; i++)
            {
                try  //send request to Omni for the current package
                {
                    CustomReadBufferRequest reqBuffer = new CustomReadBufferRequest(65, 1, 9402, i);
                    CustomReadBufferResponse packet = master.ExecuteCustomMessage<CustomReadBufferResponse>(reqBuffer);
                    //stop if find end of file
                    endIndex = packet.StrData.IndexOf(Convert.ToChar(26));//get index of end of file if exists
                    if (endIndex != -1) //if end of file was found
                    {
                        txt.Write(packet.StrData.Substring(0,endIndex));
                        break;
                    }

                    //stop if find time before yesterday at 17:00
                    try
                    {
                        CRIndex = packet.StrData.IndexOf(Convert.ToChar(13));//get index of carriage return to find date
                        if (getAlarms)
                            timeStamp = Convert.ToDateTime(packet.StrData.Substring(CRIndex + 1, 19), dateFormat);
                        else
                            timeStamp = Convert.ToDateTime(packet.StrData.Substring(CRIndex + 10, 18), dateFormat);

                        if (timeStamp < DateTime.Today.AddDays(-1).AddHours(17))
                        {
                            txt.Write(packet.StrData.Substring(0, CRIndex));
                            break;
                        }
                    }
                    catch { }
                    txt.Write(packet.StrData);
                    progress.Increment(1);
                }
                catch  //if request timeout, omni is refilling buffer
                {
                    System.Threading.Thread.Sleep(3000);//delay 3s
                    //Restart client connection
                    client.Close();
                    client = new TcpClient(OmniIP, 502);
                    client.ReceiveTimeout = 500;
                    master = ModbusIpMaster.CreateIp(client);
                    progress.Value = 0;
                    i--;//decrement packet index to reapeat the request
                }
            }

            progress.Value=63;
            txt.Close();

            client.Close();

        }
    }
}
