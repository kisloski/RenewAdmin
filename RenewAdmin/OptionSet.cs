using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;


namespace RenewAdmin
{
    public class OptionSet
    {
        //OptionSet properties
        public bool domainEnv { get; set; }
        public string userName { get; set; }
        public string groupName { get; set; }
        public string reportPath { get; set; }
        public int state { get; set; }
        public string cPath { get; set; }

        public static OptionSet configParse(string configPath)
        {
            //Parser method - gets properties values from config.xml
            OptionSet result = new OptionSet();
            XmlDocument config = new XmlDocument();

            config.Load(configPath);
            XmlNode node = config.DocumentElement.SelectSingleNode("/config/domainEnvironment");
            result.domainEnv = Convert.ToBoolean(node.InnerText);
            node = config.DocumentElement.SelectSingleNode("/config/userName");
            result.userName = node.InnerText;
            node = config.DocumentElement.SelectSingleNode("/config/groupName");
            result.groupName = node.InnerText;
            node = config.DocumentElement.SelectSingleNode("/config/reportPath");
            result.reportPath = node.InnerText;
            node = config.DocumentElement.SelectSingleNode("/config/state");
            result.state = Convert.ToInt16(node.InnerText);
            result.cPath = Program._args[0].ToString();
            //Creating report XML
            if (!File.Exists(result.reportPath))
            {
                using (XmlWriter writer = XmlWriter.Create(result.reportPath))
                {
                    writer.WriteStartElement("Report");
                    writer.WriteEndElement();
                    writer.Flush();
                }
                XmlDocument report = new XmlDocument();
                report.Load(result.reportPath);
                XmlNode node2 = report.DocumentElement.SelectSingleNode("/Report");
                XmlElement xHistory = report.CreateElement("History");
                node2.AppendChild(xHistory);
                XmlElement xReboots = report.CreateElement("Reboots");
                node2.AppendChild(xReboots);
                XmlElement xResults = report.CreateElement("Results");
                node2.AppendChild(xResults);
                report.Save(result.reportPath);
            }
            return result;
        }
        public void goToState(int newState)
        {   
            //State changing method
            XmlDocument toSave = new XmlDocument();
            toSave.Load(Program._Config.cPath);
            XmlNode state = toSave.DocumentElement.SelectSingleNode("/config/state");
            state.InnerText = newState.ToString();
            toSave.Save(Program._Config.cPath);
        }
        #region Session Variables
        public void recordSessionVars(string file, string value)
        {
            //Recoring session variables to config.xml
            if (file == "gpt")
            {
                XmlDocument config = new XmlDocument();
                config.Load(Program._Config.cPath);
                XmlNode root = config.DocumentElement.SelectSingleNode("/config");
                XmlElement xGPT = config.CreateElement("newGptLocation");
                xGPT.InnerText = value;
                root.AppendChild(xGPT);
                config.Save(Program._Config.cPath);
            }
            else if (file == "scripts")
            {
                XmlDocument config = new XmlDocument();
                config.Load(Program._Config.cPath);
                XmlNode root = config.DocumentElement.SelectSingleNode("/config");
                XmlElement xScript = config.CreateElement("newScriptsLocation");
                xScript.InnerText = value;
                root.AppendChild(xScript);
                config.Save(Program._Config.cPath);
            }
        }

        public string retrieveSessionVars(string file)
        {
            //Retrieving session variables from config.xml
            string result = "";
            XmlDocument config = new XmlDocument();
            config.Load(Program._Config.cPath);
            XmlNode root = config.DocumentElement.SelectSingleNode("/config");
            XmlNode xElem = null;
            if (file == "gpt")
            {
                xElem = root.SelectSingleNode("/config/newGptLocation");
                result = xElem.InnerText;
            }
            else if (file == "scripts" )
            {
                xElem = root.SelectSingleNode("/config/newScriptsLocation");
                result = xElem.InnerText;
            }
            return result;
        }
        #endregion

        #region Logging methods
        public void LogHistory(int state, string msg)
        {
            //Log history element method
            XmlDocument report = new XmlDocument();
            report.Load(Program._Config.reportPath);
            XmlNode root = report.DocumentElement.SelectSingleNode("/Report/History");
            XmlElement r = report.CreateElement("H");
            r.SetAttribute("ID", (root.ChildNodes.Count + 1).ToString());
            r.SetAttribute("Date", DateTime.Now.ToString());
            r.SetAttribute("State", state.ToString());
            r.SetAttribute("Val", msg);
            root.AppendChild(r);
            report.Save(Program._Config.reportPath);
        }

        public void LogReboots(int state, int rType)
        {
            //Log reboot element method
            XmlDocument report = new XmlDocument();
            report.Load(Program._Config.reportPath);
            XmlNode root = report.DocumentElement.SelectSingleNode("/Report/Reboots");
            XmlElement r = report.CreateElement("R");
            r.SetAttribute("ID", (root.ChildNodes.Count + 1).ToString());
            r.SetAttribute("Date", DateTime.Now.ToString());
            r.SetAttribute("Type", rType.ToString());
            r.SetAttribute("State", state.ToString());
            root.AppendChild(r);
            report.Save(Program._Config.reportPath);
        }

        public int GetSecondaryReboots()
        {
            //Get secondary reboots
            int count=0;
            XmlDocument report = new XmlDocument();
            report.Load(Program._Config.reportPath);
            count = report.DocumentElement.SelectNodes("/Report/Reboots/R[@Type='2']").Count;            
            return count;
            
        }

        public void LogResults()
        {
            //Create result part of report.xml
            XmlDocument report = new XmlDocument();
            report.Load(Program._Config.reportPath);
            XmlNode root = report.DocumentElement.SelectSingleNode("/Report/Results");
            PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
            GroupPrincipal gPrin = new GroupPrincipal(ctx);
            gPrin.Name = "*";
            PrincipalSearcher ps = new PrincipalSearcher();
            ps.QueryFilter = gPrin;
            PrincipalSearchResult<Principal> results = ps.FindAll();
            foreach (Principal p in results)
            {
                using (GroupPrincipal gp = (GroupPrincipal)p)
                {
                    XmlElement g = report.CreateElement("Group");
                    g.SetAttribute("Name", gp.Name);
                    root.AppendChild(g);
                    foreach (Principal p1 in gp.Members)
                    { 
                        root = report.DocumentElement.SelectSingleNode(String.Format("//Report//Results//Group[@Name='{0}']",gp.Name));
                        XmlElement u = report.CreateElement("Member");
                        u.SetAttribute("Name", p1.Name);
                        root.AppendChild(u);
                        root = report.DocumentElement.SelectSingleNode("/Report/Results");
                    }
                }
            }
            report.Save(Program._Config.reportPath);
        }
        #endregion
    }
}
