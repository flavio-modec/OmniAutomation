using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OmniAutomation
{
    class Options
    {
        private string _reportsPath;
        private string _xml004Path;
        private string _unitCode;
        private string _genTime;
        private ushort _nOmnis;
        private ushort _nOilOmnis;
        private ushort[] _fiscalOmnis;

        XmlDocument xmlOptions = new XmlDocument();

        public void start(string myDir)
        {
            xmlOptions.Load(myDir + @"\options.xml");
            _reportsPath = xmlOptions.GetElementsByTagName("Reports_Path").Item(0).InnerText;
            _xml004Path = xmlOptions.GetElementsByTagName("XML004_Path").Item(0).InnerText;
            _unitCode = xmlOptions.GetElementsByTagName("Unit_Code").Item(0).InnerText;
            _genTime = xmlOptions.GetElementsByTagName("Generation_Time").Item(0).InnerText;
            _nOmnis = Convert.ToUInt16(xmlOptions.GetElementsByTagName("Number_of_Omnis").Item(0).InnerText);
            _nOilOmnis = Convert.ToUInt16(xmlOptions.GetElementsByTagName("Number_Oil_Omnis").Item(0).InnerText);
            
            XmlNodeList fiscalOmniList = xmlOptions.GetElementsByTagName("Fiscal_Omnis").Item(0).ChildNodes;
            _fiscalOmnis = new ushort[fiscalOmniList.Count];
            for (int i = 0; i < fiscalOmniList.Count; i++ )
            {
                _fiscalOmnis[i] = Convert.ToUInt16(fiscalOmniList.Item(i).InnerText);
            }
        }

        public void save(string myDir)
        {
            xmlOptions.Load(myDir + @"\options.xml");
            xmlOptions.GetElementsByTagName("Reports_Path").Item(0).InnerText = _reportsPath;
            xmlOptions.GetElementsByTagName("XML004_Path").Item(0).InnerText = _xml004Path;
            xmlOptions.GetElementsByTagName("Unit_Code").Item(0).InnerText = _unitCode;
            xmlOptions.GetElementsByTagName("Generation_Time").Item(0).InnerText = _genTime;
            xmlOptions.GetElementsByTagName("Number_of_Omnis").Item(0).InnerText = _nOmnis.ToString();
            xmlOptions.GetElementsByTagName("Number_Oil_Omnis").Item(0).InnerText = _nOilOmnis.ToString();

            XmlNode parent = xmlOptions.GetElementsByTagName("Fiscal_Omnis").Item(0);
            XmlNode child;
            parent.RemoveAll();
            foreach (ushort FC in _fiscalOmnis)
            {
                child = parent.AppendChild(xmlOptions.CreateElement("FC"));
                child.AppendChild(xmlOptions.CreateTextNode(FC.ToString()));
            }
            xmlOptions.Save(myDir + @"\options.xml");
        }

        public string reportPath { get { return _reportsPath; } set { _reportsPath = value; } }
        public string xml004Path { get { return _xml004Path; } set { _xml004Path = value; } }
        public string unitCode { get { return _unitCode; } set { _unitCode = value; } }
        public string genTime { get { return _genTime; } set { _genTime = value; } }
        public ushort nOmnis { get { return _nOmnis; } }
        public ushort nOilOmnis { get { return _nOilOmnis; } }
        public ushort[] fiscalOmnis { get { return _fiscalOmnis; } }
    }
}
