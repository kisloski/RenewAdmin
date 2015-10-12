using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;


namespace RenewAdmin
{
    class Program
    {
        #region Global Variables
        public static OptionSet _Config = new OptionSet();
        public static string _batchPath = "";
        public static string _scriptPath = "";
        public static string _newScriptPath = "";
        public static string _gptPath = "";
        public static string _newGptPath = "";
        public static string _appPath = AppDomain.CurrentDomain.BaseDirectory.ToString();
        public static string[] _args = null;
        public static string _systemDrive = "";
        #endregion
        static void Main(string[] args)
        {
            #region Variable setup
            _systemDrive = Path.GetPathRoot(Environment.SystemDirectory);
            _batchPath = _systemDrive + "Windows\\System32\\GroupPolicy\\Machine\\Scripts\\Startup\\RenewAdmin.bat";
            //_batchPath = _systemDrive + "SYF\\renewAdminTest.bat";
            _scriptPath = _systemDrive + "Windows\\System32\\GroupPolicy\\Machine\\Scripts\\scripts.ini";
            //_scriptPath = _systemDrive + "SYF\\scripts.ini";
            _gptPath = _systemDrive + "Windows\\System32\\GroupPolicy\\gpt.ini";
            //_gptPath = _systemDrive + "SYF\\gpt.ini";
            _args = args;
            #endregion

            _Config = OptionSet.configParse(args[0].ToString());
            _Config.LogHistory( 999, "Start. Checking state.");
            _Config.LogHistory(_Config.state, "Running state: " + _Config.state.ToString());

            if (_Config.state == 0)
            { prep(); }
            else if (_Config.state == 1)
            { execution(); }
            else if (_Config.state == 2)
            { cleanup(); }
        }

        public static void prep()
        {
            _Config.LogHistory(0, "Starting Prep() method.");
            #region Create RenewAdmin.bat
            if (!File.Exists(_batchPath))
            {
                StreamWriter w = new StreamWriter(_batchPath);
                w.WriteLine("@echo off");
                w.WriteLine(@"drive = for /f ""tokens = 1 delims =\"" %%D in (""C:\Windows"") do echo %%D  ");
                w.WriteLine(@"net localgroup """ + _Config.groupName + @""" %computername%\" + _Config.userName + " /add ");
                w.WriteLine(@"cd """ + _appPath + @"""");
                w.WriteLine(@"start """" RenewAdmin.exe " + _Config.cPath);
                w.Close();
            }
            #endregion


            #region Editing scripts.ini
            Guid g = Guid.NewGuid();
            _newScriptPath = _scriptPath + ".old" + g.ToString();
            File.Copy(_scriptPath, _newScriptPath);
            _Config.LogHistory(0, "Old scripts.ini file backed up as: " + _newScriptPath.ToString());
            _Config.recordSessionVars("scripts", _newScriptPath);
            int param = 0;
            string line;
            StreamReader reader = new StreamReader(_scriptPath);
            while ((line = reader.ReadLine()) != null)
            {
                if (line.Contains("CmdLine"))
                    param++;
            }
            reader.Close();

            StreamWriter wr = new StreamWriter(_scriptPath, true);
            if (param == 0)
            {
                wr.WriteLine("");
                wr.WriteLine("[Startup]");
            }
            wr.WriteLine(param.ToString() + "CmdLine="+_batchPath);
            wr.WriteLine(param.ToString() + "Parameters=");
            wr.Close();
            #endregion


            #region Editing gpt.ini
            int vers = 0;
            Guid gu = Guid.NewGuid();
            _newGptPath = _gptPath + ".old" + gu.ToString();
            _Config.LogHistory(0, "Old gpt.ini file backed up as: " + _newGptPath.ToString());
            _Config.recordSessionVars("gpt", _newGptPath);
            File.Copy(_gptPath, _newGptPath);
            StreamReader gpt_reader = new StreamReader(_newGptPath);
            StreamWriter gpt_writer = new StreamWriter(_gptPath);
            while ((line = gpt_reader.ReadLine()) != null)
            {
                if (line.Contains("Version="))
                {
                    vers = Convert.ToInt16(line.Substring(8));
                    vers++;
                    line = "Version=" + vers.ToString();   
                }
                gpt_writer.WriteLine(line);
            }
            gpt_reader.Close();
            gpt_writer.Close();
            #endregion


            #region Run gpupdate
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C gpupdate";
            process.StartInfo = startInfo;
            process.Start();
            #endregion


            forceReboot(_Config.state, 1);
            _Config.goToState(1);
        }
        public static void execution()
        {
            _Config.LogHistory(_Config.state, "Starting execution() method.");
            PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, _Config.userName);
            GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, _Config.groupName);
            if (user.IsMemberOf(group))
            {
                _Config.LogHistory(_Config.state, String.Format("User: {0} regained membership in group {1}", _Config.userName, _Config.groupName));
                _Config.LogResults();
                _Config.goToState(2);
                forceReboot(_Config.state, 1);
            }
            else
            {
                if (_Config.GetSecondaryReboots() < 3)
                {
                    _Config.LogHistory(_Config.state, String.Format("User: {0} didn't regain membership in Group: {1}. Initializing secondary reboot {2}.", _Config.userName, _Config.groupName, (_Config.GetSecondaryReboots() + 1).ToString()));
                    forceReboot(_Config.state, 2);
                }
                else
                {
                    _Config.LogHistory(_Config.state, String.Format("User: {0} didn't regain membership in Group: {1}. Secondary reboots failed", _Config.userName, _Config.groupName));
                    //_Config.goToState(2);
                    _Config.LogResults();
                    forceReboot(_Config.state, 1);
                }
            }
        }
        public static void cleanup()
        {
            //Retrieve session variables
            _newGptPath = _Config.retrieveSessionVars("gpt");
            _newScriptPath = _Config.retrieveSessionVars("scripts");

            #region Create cleanup.bat
            _Config.LogHistory(_Config.state, "Starting cleanup() method.");
            StreamWriter cleanup = new StreamWriter(_appPath + "\\cleanup.bat");
            cleanup.WriteLine(@"@Echo off");
            cleanup.WriteLine(@"%~d1");
            cleanup.WriteLine(@"cd ""%~p1""");
            cleanup.WriteLine(@"ping 127.0.0.1 -n 30 -w 1000 > NUL");
            cleanup.WriteLine(@"DEL starter.bat");
            cleanup.WriteLine(@"DEL RenewAdmin.exe");
            cleanup.WriteLine(@"DEL config.xml");
            cleanup.WriteLine(@"DEL cleanup.bat");
            cleanup.Close();
            #endregion

            #region Restore scripts.ini
            string newpath = _newScriptPath.Substring(0, _newScriptPath.Length - 40);
            try
            {
                File.Copy(newpath, _scriptPath, true);
            }
            catch (Exception e)
            {
                _Config.LogHistory(_Config.state, e.Message);
            }
            #endregion

            #region Restore gpt.ini
            int vers = 0;
            string line = "";
            try
            {
                StreamReader gpt_reader = new StreamReader(_newGptPath);
                StreamWriter gpt_writer = new StreamWriter(_gptPath);
                while ((line = gpt_reader.ReadLine()) != null)
                {
                    if (line.Contains("Version="))
                    {
                        vers = Convert.ToInt16(line.Substring(8));
                        vers++;
                        vers++;
                        line = "Version=" + vers.ToString();
                    }
                    gpt_writer.WriteLine(line);
                }
                gpt_reader.Close();
                gpt_writer.Close();
            }
            catch (Exception e)
            {
                _Config.LogHistory(_Config.state, e.Message);
            }

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C gpupdate";
            process.StartInfo = startInfo;
            process.Start();
            #endregion

            System.Diagnostics.Process.Start("cleaner.bat");
        }
        public static void forceReboot(int state, int rType)
        {
            _Config.LogReboots(state, rType);
            System.Diagnostics.Process.Start("shutdown.exe", "-r -t 60");
        }
    }
}
