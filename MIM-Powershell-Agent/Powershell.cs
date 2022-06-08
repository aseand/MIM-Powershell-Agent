using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;

using Microsoft.MetadirectoryServices;
using Microsoft.Win32;

using NLog;

namespace MIM
{
    class PowerShellInstance
    {
        private string MA_NAME { get; set; }
        private NLog.Logger logger { get; set; }
        private PowerShell PowerShell { get; set; }
        private Dictionary<string, FileInfo> ScriptList { get; set; }

        public PowerShellInstance(string LogName, string MA_NAME)
        {
            this.logger = NLog.LogManager.GetLogger(LogName);
            this.MA_NAME = MA_NAME;

            try
            {
                this.PowerShell = null;
                this.ScriptList = new Dictionary<string, FileInfo>();
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                logger.Error(e.Source);
                logger.Error(e.StackTrace);
                throw e;
            }
        }

        public void Dispose()
        {
            this.MA_NAME = null;
            //this.Config = null;
            this.logger.Factory.Flush();
            this.logger.Factory.Dispose();
            this.logger = null;
            this.PowerShell.Dispose();
            this.ScriptList.Clear();
            this.ScriptList = null;
        }

        public Collection<PSObject> InvokeCommand(string Command, Dictionary<string, object> Parameters)
        {
            DateTime StartTime = DateTime.Now;

            Collection<PSObject> PSObjects = null;
            try
            {
                if (this.logger.IsDebugEnabled)
                {
                    this.logger.Debug($"Invoke command: {Command} Parameters: { (Parameters == null ? "" : string.Join(",", Parameters.Keys))}");
                }
                ReloadScript();

                this.PowerShell.AddCommand(Command);

                if (Parameters != null)
                {
                    this.PowerShell.AddParameters(Parameters);
                }

                PSObjects = PowerShell.Invoke();
                PowerShell.Commands.Clear();
            }
            catch(Exception e)
            {
                logger.Error(e.Message);
                logger.Error(e.Source);
                logger.Error(e.StackTrace);
                throw e;
            }

            if (this.logger.IsTraceEnabled)
                this.logger.Trace($"Invoke command: {Command} in {(DateTime.Now - StartTime).TotalMilliseconds}ms");

            return PSObjects;
        }

        void Error_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<ErrorRecord> ErrorRecords = (PSDataCollection<ErrorRecord>)sender;
            ErrorRecord ErrorRecord = ErrorRecords[e.Index];
            if (ErrorRecord != null)
            {
                this.logger.Error(ErrorRecord.Exception.Message);
                this.logger.Error(ErrorRecord.Exception.Source);
                this.logger.Error(ErrorRecord.Exception.StackTrace);
                this.logger.Error(ErrorRecord.ScriptStackTrace);
                ErrorRecords.RemoveAt(e.Index);
            }
        }

        private void ReloadScript()
        {
            bool Reload = false;
            foreach (string scriptfile in this.ScriptList.Keys)
            {
                System.IO.FileInfo newFilinfo = new System.IO.FileInfo(scriptfile);
                if (this.ScriptList[scriptfile].LastWriteTime != newFilinfo.LastWriteTime)
                {
                    this.logger.Info($"LastWriteTime change on: {scriptfile}");
                    if (this.logger.IsDebugEnabled)
                    {
                        this.logger.Debug($"{this.ScriptList[scriptfile].LastWriteTime} : {newFilinfo.LastWriteTime}");
                    }
                    Reload = true;
                    break;
                }
            }

            if (Reload)
            {
                DateTime StartTime = DateTime.Now;

                if (this.logger.IsDebugEnabled)
                    this.logger.Debug("Reload script");

                //Copy and clear list
                string[] ScriptList = new string[this.ScriptList.Keys.Count];
                this.ScriptList.Keys.CopyTo(ScriptList, 0);
                this.ScriptList.Clear();

                DateTime StartPSVariableTime = DateTime.Now;
                //Get PowerShellInstance PSVariable
                Dictionary<string, PSVariable> PSVariableList = new Dictionary<string, PSVariable>();
                Collection<PSObject> list = this.InvokeCommand("Get-Variable", null);
                foreach (PSObject item in list)
                {
                    if (item != null && item.BaseObject != null)
                    {
                        if (item.BaseObject.GetType() == typeof(PSVariable))
                        {
                            PSVariable temp = (PSVariable)item.BaseObject;
                            PSVariableList.Add(temp.Name, temp);

                            if (this.logger.IsDebugEnabled)
                                this.logger.Debug($"Save variabel '{temp.Name}'");

                        }
                    }
                }

                if (this.logger.IsDebugEnabled)
                    this.logger.Debug($"Save variabels in {(DateTime.Now - StartPSVariableTime).TotalMilliseconds}ms");


                //Dispose and Create new, BUT Dont run Initialize!
                this.PowerShell.Dispose();
                this.PowerShell = null;
                LoadScriptList(ScriptList, false);

                StartPSVariableTime = DateTime.Now;
                //Reset PSVariable
                list = this.InvokeCommand("Get-Variable", null);
                foreach (PSObject item in list)
                {
                    if (item != null && item.BaseObject != null)
                    {
                        if (item.BaseObject.GetType() == typeof(PSVariable))
                        {
                            PSVariable temp = (PSVariable)item.BaseObject;
                            //Remove all variabels that is create from new script
                            PSVariableList.Remove(temp.Name);
                        }
                    }
                }
                foreach (PSVariable item in PSVariableList.Values)
                    this.PowerShell.Runspace.SessionStateProxy.PSVariable.Set(item);
                if (this.logger.IsTraceEnabled)
                {
                    this.logger.Trace($"Set PSVariable(s): {string.Join(", ", PSVariableList.Keys)} in {(DateTime.Now - StartPSVariableTime).TotalMilliseconds}ms");
                    this.logger.Trace($"Reload script in {(DateTime.Now - StartTime).TotalMilliseconds}ms");
                }
            }

        }

        public void LoadScriptList(string Script, bool RunInitialize)
        {
            LoadScriptList(new string[] { Script }, RunInitialize);
        }

        public void LoadScriptList(string[] ScriptList, bool RunInitialize)
        {
            DateTime StartTime = DateTime.Now;

            if (this.PowerShell == null)
            {
                if (this.logger.IsDebugEnabled)
                    this.logger.Debug("Create PowerShellInstance");

                this.PowerShell = PowerShell.Create();
                this.PowerShell.Streams.Error.DataAdded += this.Error_DataAdded;
            }

            if (this.logger.IsDebugEnabled)
                this.logger.Debug("Load Script(s)");

            foreach (string scriptfile in ScriptList)
            {
                if (System.IO.File.Exists(scriptfile))
                {
                    
                    if (!this.ScriptList.ContainsKey(scriptfile))
                    {
                        DateTime StartLoadTime = DateTime.Now;
                        if (this.logger.IsDebugEnabled)
                            this.logger.Debug($"Load: {scriptfile}");

                        //load scriptfile
                        this.PowerShell.AddScript($". '{scriptfile}'");
                        this.PowerShell.Invoke();
                        this.PowerShell.Commands.Clear();

                        //Add fileinfo to list
                        System.IO.FileInfo fileinfo = new System.IO.FileInfo(scriptfile);
                        this.ScriptList.Add(scriptfile, fileinfo);

                        if (this.logger.IsTraceEnabled)
                            this.logger.Trace($"Loaded {scriptfile} in : {(DateTime.Now - StartLoadTime).TotalMilliseconds}ms");
                    }

                }
                else
                {
                    this.logger.Error($"{scriptfile} dont exist");
                }
            }



            if (RunInitialize)
            {
                DateTime StartInitializeTime = DateTime.Now;

                if (this.logger.IsDebugEnabled)
                    this.logger.Debug("Initialize script");
                InvokeCommand("Initialize", new Dictionary<string, object>() { { "logger", this.logger }, { "MAName", this.MA_NAME } });

                if (this.logger.IsTraceEnabled)
                    this.logger.Trace($"Initialize script in {(DateTime.Now - StartInitializeTime).TotalMilliseconds}ms");
            }

            if (this.logger.IsTraceEnabled)
                this.logger.Trace($"Create PowerShellInstance in {(DateTime.Now - StartTime).TotalMilliseconds}ms");
        }

    }

    public class Powershell :
        IMASynchronization,
        IMVSynchronization,
        IMAExtensible2GetCapabilitiesEx,
        IMAExtensible2GetSchema,
        IMAExtensible2GetParameters,
        IMAExtensible2CallImport,
        IMAExtensible2CallExport,
        IMAExtensible2Password
    {
        private NLog.Logger logger = null;
        private string loggerFullName;
        private Dictionary<string, PowerShellInstance> PowerShellInstance = new Dictionary<string, PowerShellInstance>();
        private string[] IMVSynchronizationList;
        private string IMAExtensible2InitalScript = null;

        public Powershell()
        {
            RegistryKey Registry = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey regKey = Registry.OpenSubKey(@"Software\aseand\MIM-Powershell-Agent");

            Dictionary<string, string> RegConfig = new Dictionary<string, string>();
            foreach (string name in regKey.GetValueNames())
            {
                RegConfig.Add(name, regKey.GetValue(name).ToString());
            }
            regKey.Close();


            string LoggingConfigurationFullPath = RegConfig.ContainsKey("LoggingConfigurationFullPath") ? RegConfig["LoggingConfigurationFullPath"] : "";
            IMAExtensible2InitalScript = RegConfig.ContainsKey("IMAExtensible2InitalScript") ? RegConfig["IMAExtensible2InitalScript"] : "";

            string LoggingPath = "";
            try
            {
                LoggingPath = MAUtils.MAFolder;
            }
            catch
            {
                try
                {
                    LoggingPath = Microsoft.MetadirectoryServices.Utils.WorkingDirectory;
                }
                catch
                {
                    LoggingPath = RegConfig.ContainsKey("LoggingPath") ? RegConfig["LoggingPath"] : Path.GetTempPath();
                }
            }
             
            int ArchiveAboveSizeMb = RegConfig.ContainsKey("ArchiveAboveSizeMb") ? int.Parse(RegConfig["ArchiveAboveSizeMb"]) : 10;
            int MaxArchiveFiles = RegConfig.ContainsKey("MaxArchiveFiles") ? int.Parse(RegConfig["MaxArchiveFiles"]) : 20;

            LogLevel MaxLogLevel = RegConfig.ContainsKey("MaxLogLevel") ? LogLevel.FromString(RegConfig["MaxLogLevel"]) : LogLevel.Fatal;
            LogLevel MinLogLevel = RegConfig.ContainsKey("MinLogLevel") ? LogLevel.FromString(RegConfig["MinLogLevel"]) : LogLevel.Info;

            loggerFullName = typeof(Powershell).FullName;

            if (!string.IsNullOrEmpty(LoggingConfigurationFullPath) && File.Exists(LoggingConfigurationFullPath))
            {
                NLog.LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration(LoggingConfigurationFullPath);
            }
            else
            {
                NLog.Targets.FileTarget target = new NLog.Targets.FileTarget()
                {
                    Name = loggerFullName,
                    FileName = Path.Combine(LoggingPath, "MIM-Powershell-Agent.log"),
                    ArchiveAboveSize = ArchiveAboveSizeMb * 1048576,
                    MaxArchiveFiles = MaxArchiveFiles,
                    ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.Rolling,
                    Encoding = Encoding.UTF8
                };

                NLog.Targets.Wrappers.AsyncTargetWrapper asyncTargetWrapper = new NLog.Targets.Wrappers.AsyncTargetWrapper()
                {
                    WrappedTarget = target
                };

                NLog.Config.LoggingConfiguration LoggingConfiguration = new NLog.Config.LoggingConfiguration();
                LoggingConfiguration.AddRule(MinLogLevel, MaxLogLevel, asyncTargetWrapper);
                NLog.LogManager.Configuration = LoggingConfiguration;
            }
            
            logger = NLog.LogManager.GetLogger(loggerFullName);

            if (logger.IsDebugEnabled)
            {
                logger.Debug($"{loggerFullName} Initialize");
            }
        }

        public void Dispose()
        {
            this.logger.Factory.Flush();
            this.logger.Factory.Dispose();
            this.logger = null;
        }
        /// <summary>
        /// Get first object in Collection that match type
        /// Object is cast as selected
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="PSCollection"></param>
        /// <returns></returns>
        private T GetFirstObjectOf<T>(Collection<PSObject> PSCollection)
        {
            T returnObject = default(T);
            foreach (PSObject obj in PSCollection)
            {
                if (obj != null && obj.BaseObject != null)
                {
                    if (logger.IsDebugEnabled)
                        logger.Debug($"Object type : {obj.BaseObject.GetType().Name}");

                    if (obj.BaseObject.GetType() == typeof(T))
                    {
                        returnObject = (T)obj.BaseObject;
                        break;
                    }
                }
                else if (logger.IsDebugEnabled)
                    logger.Debug("Object is null!");
            }
            return returnObject;
        }

        #region IMASynchronization

        void IMASynchronization.Initialize()
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.Initialize {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
            }
            string IMASynchronizationPath = Path.Combine(Microsoft.MetadirectoryServices.Utils.WorkingDirectory, "IMASynchronization.ps1");
            if (logger.IsDebugEnabled)
                logger.Debug($"IMASynchronization path: {IMASynchronizationPath}");

            CurentPowerShellInstance.LoadScriptList(IMASynchronizationPath, true);
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.Initialize", null);

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.Initialize {curent_MA_NAME}");
        }

        void IMASynchronization.Terminate()
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.Terminate {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.Terminate", null);

            logger.Factory.Flush();

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.Terminate {curent_MA_NAME}");
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.ShouldProjectToMV {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }

            Collection<PSObject> List = CurentPowerShellInstance.InvokeCommand("IMASynchronization.ShouldProjectToMV", new Dictionary<string, object>() { { "CSEntry", csentry } });
            bool result = GetFirstObjectOf<bool>(List);
            MVObjectType = GetFirstObjectOf<string>(List);


            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.ShouldProjectToMV {curent_MA_NAME}");

            return result;
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.Deprovision {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }

            DeprovisionAction result = GetFirstObjectOf<DeprovisionAction>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.Deprovision", new Dictionary<string, object>() { { "CSEntry", csentry } }));

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.Deprovision {curent_MA_NAME}");

            return result;
        }

        bool IMASynchronization.FilterForDisconnection(CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.FilterForDisconnection {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            bool result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.FilterForDisconnection", new Dictionary<string, object>() { { "CSEntry", csentry } }));

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.FilterForDisconnection {curent_MA_NAME}");

            return result;
        }

        void IMASynchronization.MapAttributesForJoin(string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
            {
                logger.Debug($"Start IMASynchronization.MapAttributesForJoin {curent_MA_NAME}");
                logger.Debug($"FlowRuleName: {FlowRuleName} ");
            }

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.MapAttributesForJoin", new Dictionary<string, object>() { 
            { "FlowRuleName", FlowRuleName } ,
            { "CSEntry", csentry } ,
            { "ValueCollection", values } 
            });


            if (logger.IsDebugEnabled)
            {
                logger.Debug($"ValueCollection: {(string.Join(", ", values.ToStringArray()))}");
                logger.Debug($"Done IMASynchronization.MapAttributesForJoin {curent_MA_NAME}");
            }
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.ResolveJoinSearch {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            Collection<PSObject> List = CurentPowerShellInstance.InvokeCommand("IMASynchronization.ResolveJoinSearch", new Dictionary<string, object>() { 
            { "joinCriteriaName", joinCriteriaName } ,
            { "CSEntry", csentry } ,
            { "rgmventry", rgmventry } 
            });

            bool result = GetFirstObjectOf<bool>(List);
            imventry = GetFirstObjectOf<int>(List);

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.ResolveJoinSearch {curent_MA_NAME}");

            return result;
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMASynchronization.MapAttributesForImport {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            try
            {
                bool result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.MapAttributesForImport", new Dictionary<string, object>() { 
                { "FlowRuleName", FlowRuleName } ,
                { "CSEntry", csentry } ,
                { "MVEntry", mventry } 
                }));

            }
            catch (Exception e)
            {
                if (e.Message == "Microsoft.MetadirectoryServices.DeclineMappingException")
                {

                    throw new Microsoft.MetadirectoryServices.DeclineMappingException();
                }
                else
                {
                    logger.Error(curent_MA_NAME);
                    logger.Error(e.GetType());
                    logger.Error(e.Message);
                    logger.Error(e.Source);
                    logger.Error(e.StackTrace);
                    throw(e);
                }
            }


            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.MapAttributesForImport {curent_MA_NAME}");
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForExport {0}", curent_MA_NAME);


            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            bool result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.MapAttributesForExport", new Dictionary<string, object>() { 
            { "FlowRuleName", FlowRuleName } ,
            { "CSEntry", csentry } ,
            { "MVEntry", mventry } 
            }));

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMASynchronization.MapAttributesForExport {curent_MA_NAME}");
        }

        #endregion

        #region IMVSynchronization

        void IMVSynchronization.Initialize()
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Initialize");

            List<string> MAIMVSynchronizationList = new List<string>();
            var MAs = Microsoft.MetadirectoryServices.Utils.MAs.GetEnumerator();
            while (MAs.MoveNext())
            {
                string IMVSynchronizationPath = Path.Combine(Microsoft.MetadirectoryServices.Utils.WorkingDirectory, MAs.Current, "IMVSynchronization.ps1");

                if (File.Exists(IMVSynchronizationPath))
                {
                    MAIMVSynchronizationList.Add(MAs.Current);

                    if (logger.IsDebugEnabled)
                        logger.Debug($"Found IMASynchronization script path: {IMVSynchronizationPath}");

                    PowerShellInstance CurentPowerShellInstance;
                    if (!PowerShellInstance.TryGetValue(MAs.Current, out CurentPowerShellInstance))
                    {
                        CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{MAs.Current}", MAs.Current);
                        PowerShellInstance.Add(MAs.Current, CurentPowerShellInstance);

                        CurentPowerShellInstance.LoadScriptList(IMVSynchronizationPath, true);
                        CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Initialize", null);
                    }

                    if (logger.IsDebugEnabled)
                        logger.Debug($"Done IMVSynchronization.Initialize {MAs.Current}");
                }
            }
            this.IMVSynchronizationList = MAIMVSynchronizationList.ToArray();

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Initialize");
        }

        void IMVSynchronization.Terminate()
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Terminate");

            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                    throw (new Exception("PowerShell instance is null"));
                }

                CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Terminate", null);

            }

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Terminate");
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            DateTime StartTime = DateTime.Now;

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Provision");

            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                    throw (new Exception("PowerShell instance is null"));
                }
                DateTime StartInvokeTime = DateTime.Now;
                if (logger.IsDebugEnabled)
                    logger.Debug($"Start IMVSynchronization.Provision on {curent_MA_NAME}");
                CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Provision", new Dictionary<string, object>() { { "MVEntry", mventry } });
                if (logger.IsTraceEnabled)
                    logger.Trace($"Start IMVSynchronization.Provision on {curent_MA_NAME} in {(DateTime.Now - StartInvokeTime).TotalMilliseconds}ms");

            }

            if (logger.IsTraceEnabled)
                logger.Trace($"Done IMVSynchronization.Provision in {(DateTime.Now - StartTime).TotalMilliseconds}ms");
        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.ShouldDeleteFromMV");

            bool result = false;
            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                    throw (new Exception("PowerShell instance is null"));
                }
                result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMVSynchronization.ShouldDeleteFromMV", new Dictionary<string, object>() { 
                    { "CSEntry", csentry } ,
                    { "MVEntry", mventry }
                    }));

            }

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.ShouldDeleteFromMV");

            return result;
        }

        #endregion

        #region page setings

        public int ImportDefaultPageSize
        {
            get
            {
                return 100;
            }
        }

        public int ImportMaxPageSize
        {
            get
            {
                return 10000;
            }
        }

        int IMAExtensible2CallExport.ExportDefaultPageSize
        {
            get
            {
                return 100;
            }
        }

        int IMAExtensible2CallExport.ExportMaxPageSize
        {
            get
            {
                return 100;
            }
        }

        #endregion

        #region IMAExtensible2 setup

        System.Collections.Generic.IList<ConfigParameterDefinition> IMAExtensible2GetParameters.GetConfigParameters(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2GetParameters.GetConfigParameters {page.ToString("f")}"); 

            System.Collections.Generic.List<ConfigParameterDefinition> list = null;
            string IMAExtensible2Script = configParameters.Contains("IMAExtensible2GetParameters") ? configParameters["IMAExtensible2GetParameters"].Value : IMAExtensible2InitalScript;

            if (IMAExtensible2Script != null && File.Exists(IMAExtensible2Script))
            {
                if (logger.IsDebugEnabled)
                    logger.Debug($"Load: {IMAExtensible2Script}");

                PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}", loggerFullName);
                CurrentPowerShellInstance.LoadScriptList(IMAExtensible2Script, true);
                list = GetFirstObjectOf<System.Collections.Generic.List<ConfigParameterDefinition>>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetParameters.GetConfigParameters", new Dictionary<string, object>() {
                    { "ConfigParameters", configParameters } ,
                    { "ConfigParameterPage", page }
                    }));

                CurrentPowerShellInstance.Dispose();
                CurrentPowerShellInstance = null;
            }
            else
            {
                #region IMAExtensible2GetParameters
                try
                {
                    list = new List<ConfigParameterDefinition>();
                    switch (page)
                    {
                        case ConfigParameterPage.Capabilities:

                            //DistinguishedNameStyle         : Generic
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateDropDownParameter("DistinguishedNameStyle", new string[] { "Generic", "Ldap", "None" }, false, "Generic"));
                            //ObjectRename                   : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("ObjectRename", false));
                            //NoReferenceValuesInFirstExport : False
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("NoReferenceValuesInFirstExport", false));
                            //DeltaImport                    : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("DeltaImport", true));
                            //ConcurrentOperation            : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("ConcurrentOperation", true));
                            //DeleteAddAsReplace             : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("DeleteAddAsReplace", true));
                            //ExportPasswordInFirstPass      : False
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("ExportPasswordInFirstPass", false));
                            //FullExport                     : False
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("FullExport", false));
                            //ObjectConfirmation             : Normal
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateDropDownParameter("ObjectConfirmation", new string[] { "Normal", "NoDeleteConfirmation", "NoAddAndDeleteConfirmation" }, false, "Normal"));
                            //ExportType                     : ObjectReplace
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateDropDownParameter("ExportType", new string[] { "AttributeUpdate", "AttributeReplace", "ObjectReplace", "MultivaluedReferenceAttributeUpdate" }, false, "AttributeUpdate"));
                            //Normalizations                 : None
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateDropDownParameter("Normalizations", new string[] { "None", "Uppercase", "RemoveAccents" }, false, "None"));
                            //IsDNAsAnchor                   : False
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("IsDNAsAnchor", false));
                            //SupportImport                  : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("SupportImport", true));
                            //SupportExport                  : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("SupportExport", true));
                            //SupportPartitions              : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("SupportPartitions", false));
                            //SupportPassword                : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("SupportPassword", false));
                            //SupportHierarchy               : True
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateCheckBoxParameter("SupportHierarchy", false));
                            break;

                        case ConfigParameterPage.Connectivity:
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateLabelParameter("Powershell script:"));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateLabelParameter("(Optinal)"));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2GetParameters", "", IMAExtensible2Script));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2GetCapabilitiesEx", "", ""));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateLabelParameter("(Mandatory)"));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2GetSchema", "", IMAExtensible2Script));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2CallImport", "", IMAExtensible2Script));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2CallExport", "", IMAExtensible2Script));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateLabelParameter("(Optinal)"));
                            list.Add(Microsoft.MetadirectoryServices.ConfigParameterDefinition.CreateStringParameter("IMAExtensible2Password", "", IMAExtensible2Script));
                            break;

                        case ConfigParameterPage.Global: break;
                        case ConfigParameterPage.Partition: break;
                        case ConfigParameterPage.RunStep: break;
                        case ConfigParameterPage.Schema: break;
                    }
                }
                catch (Exception e)
                {
                    logger.Error(e.Message);
                    logger.Error(e.Source);
                    logger.Error(e.StackTrace);
                    throw e;
                }
                #endregion
            }

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.GetConfigParameters");

            return list;
        }

        MACapabilities IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters)
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");

            MACapabilities result;
            string IMAExtensible2GetCapabilitiesEx = configParameters["IMAExtensible2GetCapabilitiesEx"].Value;
            if (File.Exists(IMAExtensible2GetCapabilitiesEx))
            {
                if (logger.IsDebugEnabled)
                    logger.Debug($"Load: {IMAExtensible2GetCapabilitiesEx}");

                PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}", loggerFullName);
                CurrentPowerShellInstance.LoadScriptList(IMAExtensible2GetCapabilitiesEx,true);
                result = GetFirstObjectOf<MACapabilities>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx", null));

                CurrentPowerShellInstance.Dispose();
                CurrentPowerShellInstance = null;
            }
            else
            {
                try
                {
                    result = new MACapabilities();
                    result.DistinguishedNameStyle = (MADistinguishedNameStyle)Enum.Parse(typeof(MADistinguishedNameStyle), configParameters["DistinguishedNameStyle"].Value);
                    result.ObjectRename = configParameters["ObjectRename"].Value == "1" ? true : false;
                    result.NoReferenceValuesInFirstExport = configParameters["NoReferenceValuesInFirstExport"].Value == "1" ? true : false;
                    result.DeltaImport = configParameters["DeltaImport"].Value == "1" ? true : false;
                    result.ConcurrentOperation = configParameters["ConcurrentOperation"].Value == "1" ? true : false;
                    result.DeleteAddAsReplace = configParameters["DeleteAddAsReplace"].Value == "1" ? true : false;
                    result.ExportPasswordInFirstPass = configParameters["ExportPasswordInFirstPass"].Value == "1" ? true : false;
                    result.FullExport = configParameters["FullExport"].Value == "1" ? true : false;
                    result.ObjectConfirmation = (MAObjectConfirmation)Enum.Parse(typeof(MAObjectConfirmation), configParameters["ObjectConfirmation"].Value);
                    result.ExportType = (MAExportType)Enum.Parse(typeof(MAExportType), configParameters["ExportType"].Value);
                    result.Normalizations = (MANormalizations)Enum.Parse(typeof(MANormalizations), configParameters["Normalizations"].Value);
                    result.IsDNAsAnchor = configParameters["IsDNAsAnchor"].Value == "1" ? true : false;
                    result.SupportImport = configParameters["SupportImport"].Value == "1" ? true : false;
                    result.SupportExport = configParameters["SupportExport"].Value == "1" ? true : false;
                    result.SupportPartitions = configParameters["SupportPartitions"].Value == "1" ? true : false;
                    result.SupportPassword = configParameters["SupportPassword"].Value == "1" ? true : false;
                    result.SupportHierarchy = configParameters["SupportHierarchy"].Value == "1" ? true : false;
                }
                catch (Exception e)
                {
                    logger.Error(e.Message);
                    logger.Error(e.Source);
                    logger.Error(e.StackTrace);
                    throw new ExtensibleExtensionException(e.Message);
                }
            }

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");

            return result;
        }

        Schema IMAExtensible2GetSchema.GetSchema(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters)
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetSchema.GetSchema");

            Schema result;
            if (logger.IsDebugEnabled)
                logger.Debug($"Load: {configParameters["IMAExtensible2GetSchema"].Value}");

            PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}", loggerFullName);
            CurrentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2GetSchema"].Value,true);
            result = GetFirstObjectOf<Schema>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetSchema.GetSchema", new Dictionary<string, object>() { { "ConfigParameters", configParameters } }));

            CurrentPowerShellInstance.Dispose();
            CurrentPowerShellInstance = null;

            if (logger.IsDebugEnabled)
            {
                //log schema
                foreach (SchemaType type in result.Types)
                {
                    logger.Debug("Schema type: {0}", type.Name);
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug($"Schema attibute{attibute.Name}: {(attibute.IsMultiValued ? "(IsMultiValued)" : "")}");
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug($"Schema AnchorAttributes: { attibute.Name}");
                    }
                }
            }
            logger.Debug("Done IMAExtensible2GetSchema.GetSchema");
            return result;
        }

        ParameterValidationResult IMAExtensible2GetParameters.ValidateConfigParameters(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            if (logger.IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetParameters.ValidateConfigParameters");

            ParameterValidationResult result = new ParameterValidationResult();

            try
            {
                //Optinal
                foreach (string param in new string[] { "IMAExtensible2GetParameters", "IMAExtensible2GetCapabilitiesEx", "IMAExtensible2Password" })
                {
                    if (configParameters.Contains(param) && !string.IsNullOrEmpty(configParameters[param].Value))
                    {
                        if (!File.Exists(configParameters[param].Value))
                        {
                            result.Code = ParameterValidationResultCode.Failure;
                            result.ErrorMessage = $"File{param} don´t exist {configParameters[param].Value}";
                            result.ErrorParameter = param;
                            return result;
                        }
                    }
                }
                //Mandatory
                foreach (string param in new string[] { "IMAExtensible2GetSchema", "IMAExtensible2CallImport", "IMAExtensible2CallExport" })
                {
                    if (!configParameters.Contains(param) || string.IsNullOrEmpty(configParameters[param].Value))
                    {
                        result.Code = ParameterValidationResultCode.Failure;
                        result.ErrorMessage = $"Missing configParameters: {param}";
                        result.ErrorParameter = param;
                        return result;
                    }
                    if (!File.Exists(configParameters[param].Value))
                    {
                        result.Code = ParameterValidationResultCode.Failure;
                        result.ErrorMessage = $"File({param}) don´t exist {configParameters[param].Value}";
                        result.ErrorParameter = param;
                        return result;
                    }
                }

                if (File.Exists(configParameters["IMAExtensible2GetParameters"].Value))
                {
                    if (logger.IsDebugEnabled)
                        logger.Debug("Load: {0}", configParameters["IMAExtensible2GetParameters"].Value);

                    PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}", loggerFullName);
                    CurrentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2GetParameters"].Value, true);
                    result = GetFirstObjectOf<ParameterValidationResult>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetParameters.ValidateConfigParameters", new Dictionary<string, object>() {
            { "ConfigParameters", configParameters } ,
            { "ConfigParameterPage", page }
            }));

                    CurrentPowerShellInstance.Dispose();
                    CurrentPowerShellInstance = null;
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                logger.Error(e.Source);
                logger.Error(e.StackTrace);
                throw e;
            }

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.ValidateConfigParameters");

            return result;
        }

        #endregion

        #region import

        OpenImportConnectionResults IMAExtensible2CallImport.OpenImportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenImportConnectionRunStep openImportRunStep)
        {
            OpenImportConnectionResults result = null;
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallImport.OpenImportConnection {0}", curent_MA_NAME);
                logger.Debug("openImportRunStep: Type: {0} PageSize: {1} CustomData: {2}", openImportRunStep.ImportType, openImportRunStep.PageSize, openImportRunStep.CustomData);

                //log schema
                foreach (SchemaType type in types.Types)
                {
                    logger.Debug("Schema type: {0}", type.Name);
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug($"Schema attibute{(attibute.IsMultiValued ? "(IsMultiValued)" : "")}: { attibute.Name}");
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug($"Schema AnchorAttributes: {attibute.Name}");
                    }
                }
            }

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
            }
            CurentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2CallImport"].Value,true);
            result = GetFirstObjectOf<OpenImportConnectionResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallImport.OpenImportConnection", new Dictionary<string, object>() { 
                { "ConfigParameters", configParameters } ,
                { "Schema", types } ,
                { "OpenImportConnectionRunStep", openImportRunStep } 
                }));

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.OpenImportConnection");

            return result;
        }

        GetImportEntriesResults IMAExtensible2CallImport.GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallImport.GetImportEntries {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            GetImportEntriesResults result = GetFirstObjectOf<GetImportEntriesResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallImport.GetImportEntries", new Dictionary<string, object>() { 
            { "GetImportEntriesRunStep", importRunStep }
            }));

            if (logger.IsDebugEnabled)
            {
                logger.Debug("MoreToImport? {0}", result.MoreToImport);
                logger.Debug("Done IMAExtensible2CallImport.GetImportEntries {0}", curent_MA_NAME);
            }

            return result;
        }

        CloseImportConnectionResults IMAExtensible2CallImport.CloseImportConnection(CloseImportConnectionRunStep importRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2CallImport.CloseImportConnection {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CloseImportConnectionResults result = GetFirstObjectOf<CloseImportConnectionResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallImport.CloseImportConnection", new Dictionary<string, object>() { 
            { "CloseImportConnectionRunStep", importRunStep }
            }));

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2CallImport.CloseImportConnection {curent_MA_NAME}");

            return result;
        }
        #endregion

        #region export

        void IMAExtensible2CallExport.OpenExportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenExportConnectionRunStep exportRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
            {
                logger.Debug($"Start IMAExtensible2CallExport.OpenExportConnection {curent_MA_NAME}");
                logger.Debug($"OpenExportConnection: Type: {exportRunStep.ExportType} BatchSize: {exportRunStep.BatchSize} StepPartition: {exportRunStep.StepPartition.Name}");

                //log schema
                foreach (SchemaType type in types.Types)
                {
                    logger.Debug($"Schema type: {type.Name}");
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug($"Schema attibute{(attibute.IsMultiValued ? "(IsMultiValued)" : "")}: {attibute.Name}");
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug($"Schema AnchorAttributes: {attibute.Name}");
                    }
                }
            }

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
            }
            CurentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2CallExport"].Value,true);
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallExport.OpenExportConnection", new Dictionary<string, object>() { 
                { "ConfigParameters", configParameters } ,
                { "Schema", types } ,
                { "OpenExportConnectionRunStep", exportRunStep } 
                });

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2CallExport.OpenExportConnection {curent_MA_NAME}");
        }

        PutExportEntriesResults IMAExtensible2CallExport.PutExportEntries(System.Collections.Generic.IList<CSEntryChange> csentries)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2CallExport.PutExportEntries {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            
            PutExportEntriesResults result = GetFirstObjectOf<PutExportEntriesResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallExport.PutExportEntries", new Dictionary<string, object>() { 
            { "CSEntryChanges", csentries }
            }));

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2CallExport.PutExportEntries {curent_MA_NAME}");

            return result;
        }

        void IMAExtensible2CallExport.CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2CallExport.CloseExportConnection {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallExport.CloseExportConnection", new Dictionary<string, object>() { 
            { "CloseExportConnectionRunStep", exportRunStep }
            });

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2CallExport.CloseExportConnection {curent_MA_NAME}");
        }
        #endregion

        #region password
        void IMAExtensible2Password.OpenPasswordConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Partition partition)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2Password.OpenPasswordConnection {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
            }
            CurentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2Password"].Value,true);
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.OpenPasswordConnection", new Dictionary<string, object>() { 
                { "ConfigParameters", configParameters } ,
                { "Partition", partition } 
                });

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2Password.OpenPasswordConnection {curent_MA_NAME}");
        }

        ConnectionSecurityLevel IMAExtensible2Password.GetConnectionSecurityLevel()
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.GetConnectionSecurityLevel {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            ConnectionSecurityLevel result = GetFirstObjectOf<ConnectionSecurityLevel>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.GetConnectionSecurityLevel", null));

            if (logger.IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.GetConnectionSecurityLevel {0}", curent_MA_NAME);

            return result;
        }

        void IMAExtensible2Password.ClosePasswordConnection()
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2Password.ClosePasswordConnection {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.ClosePasswordConnection", null);

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2Password.ClosePasswordConnection {curent_MA_NAME}");
        }

        void IMAExtensible2Password.ChangePassword(CSEntry csentry, System.Security.SecureString oldPassword, System.Security.SecureString newPassword)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2Password.ChangePassword {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.ChangePassword", new Dictionary<string, object>() { 
            { "CSEntry", csentry },
            { "oldPassword", oldPassword },
            { "newPassword", newPassword },
            });

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2Password.ChangePassword {curent_MA_NAME}");
        }

        void IMAExtensible2Password.SetPassword(CSEntry csentry, System.Security.SecureString newPassword, PasswordOptions options)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (logger.IsDebugEnabled)
                logger.Debug($"Start IMAExtensible2Password.SetPassword {curent_MA_NAME}");

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error($"Geting powershell instance for {curent_MA_NAME}");
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.SetPassword", new Dictionary<string, object>() { 
            { "CSEntry", csentry },
            { "newPassword", newPassword },
            { "PasswordOptions", options },
            });

            if (logger.IsDebugEnabled)
                logger.Debug($"Done IMAExtensible2Password.SetPassword {curent_MA_NAME}");
        }
        #endregion
    }
}
