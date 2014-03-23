using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace OmniAutomation
{
    class XMLcreator
    {
        XmlDocument xml004 = new XmlDocument();

        public void getXML004(ushort[] FCs, string[] SN, string codInstalacao, string directory, string reportsPath)
        {
            xml004.AppendChild((xml004.CreateXmlDeclaration("1.0", "UTF-8", null)));
            
            XmlNode listaDados = xml004.AppendChild(xml004.CreateElement("a004")).AppendChild(xml004.CreateElement("LISTA_DADOS_BASICOS"));
            XmlNode dados;
            XmlNode listaAlarmes;
            XmlNode listaEventos;
            XmlNode currentNode;
            XmlNode child;

            TextReader txt;
            string line;
            DateTime time;
            DateTime today17 = DateTime.Today.AddHours(17);
            DateTime yesterday17 = DateTime.Today.AddDays(-1).AddHours(17);
            ushort modbusID;
            IFormatProvider dateFormat = new System.Globalization.CultureInfo("en-GB");

            foreach (ushort FC in FCs)
            {
                
                dados = listaDados.AppendChild(xml004.CreateElement("DADOS_BASICOS"));
                dados.Attributes.Append(xml004.CreateAttribute("NUM_SERIE_COMPUTADOR_VAZAO")).Value = SN[FC-1];
                dados.Attributes.Append(xml004.CreateAttribute("COD_INSTALACAO")).Value = codInstalacao;

                listaAlarmes = dados.AppendChild(xml004.CreateElement("LISTA_ALARMES"));

                txt = new StreamReader (reportsPath + @"\FC" + Convert.ToString(FC) + "_alarms.txt");
                //txt = new StreamReader(@"C:\OmniReports\FC9_alarms.txt");
                for (ushort i = 0; i < 6; i++) { txt.ReadLine(); }

                for (ushort i = 0; i < 500; i++)
                {
                    line = txt.ReadLine();
                    if (string.IsNullOrEmpty(line)) break;
                    try { time = Convert.ToDateTime(line.Substring(0, 18), dateFormat); }
                    catch { continue; }
                    if (time < yesterday17) break; 
                    if (time < today17)
                    {
                        currentNode = listaAlarmes.AppendChild(xml004.CreateElement("ALARMES"));
                        child = currentNode.AppendChild(xml004.CreateElement("DSC_DADO_ALARMADO"));
                        child.AppendChild(xml004.CreateTextNode("Descrição:" + line.Substring(20,30) + "-" +line.Substring(51,5)));
                        child = currentNode.InsertAfter(xml004.CreateElement("DHA_ALARME"),child);
                        child.AppendChild(xml004.CreateTextNode(time.ToString(dateFormat)));
                        child = currentNode.InsertAfter(xml004.CreateElement("DSC_MEDIDA_ALARMADA"),child);
                        child.AppendChild(xml004.CreateTextNode("Descrição:" + line.Substring(60,12)));
                    }
                }
                txt.Close();

                listaEventos = dados.AppendChild(xml004.CreateElement("LISTA_EVENTOS"));

                txt = new StreamReader (reportsPath + @"\FC" + Convert.ToString(FC) + "_events.txt");
                //txt = new StreamReader(@"C:\OmniReports\FC9_events.txt");
                for (ushort i = 0; i < 8; i++) { txt.ReadLine(); }

                for (ushort i = 0; i < 150; i++)
                {
                    line = txt.ReadLine();
                    if (string.IsNullOrEmpty(line)) break;
                    try { time = Convert.ToDateTime(line.Substring(9, 17), dateFormat); }
                    catch { continue; }
                    modbusID = Convert.ToUInt16(line.Substring(30, 5));
                    if (time < yesterday17) { break; }
                    if (time < today17 && modbusID != 100)  //eventID=100 is "Privilege Passwd" (not register in XML)
                    {
                        currentNode = listaEventos.AppendChild(xml004.CreateElement("EVENTOS"));
                        child = currentNode.AppendChild(xml004.CreateElement("DSC_DADO_ALTERADO"));
                        child.AppendChild(xml004.CreateTextNode("Event ID: " + line.Substring(1, 6) + " Modbus ID: " + line.Substring(30, 5)));
                        child = currentNode.InsertAfter(xml004.CreateElement("DSC_CONTEUDO_ORIGINAL"), child);
                        child.AppendChild(xml004.CreateTextNode("Descrição: " + line.Substring(37, 16)));
                        child = currentNode.InsertAfter(xml004.CreateElement("DSC_CONTEUDO_ATUAL"), child);
                        child.AppendChild(xml004.CreateTextNode("Descrição: " + line.Substring(55, 16)));
                        child = currentNode.InsertAfter(xml004.CreateElement("DHA_OCORRENCIA_EVENTO"), child);
                        child.AppendChild(xml004.CreateTextNode(time.ToString(dateFormat)));
                    }
                }
                txt.Close();
            }

            xml004.Save(directory + @"\004_33000167_" + DateTime.Now.ToString("yyyyMMddHHmmss",dateFormat) + "_" + codInstalacao + ".xml");
        }
    }
}
