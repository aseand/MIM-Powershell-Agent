/* 2015-02-09 Anders Åsén
 * 2015-09-17 Add support for run ps on IMASynchronization, IMVSynchronization, IMAExtensible2
 * 2016-02-17 Add support for ReloadPowerShellScript (LastWriteTime)
 *            Add support for DeclineMappingException
 * 2016-05-19 Fix minor error
 * 2016-06-01 Add trace logs for exec time
 * 2017-04-13 Fix '.' name bug
 * 2018-09-06 Rename namespace
 *              
 */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Management.Automation;
using Microsoft.MetadirectoryServices;
using System.IO;

using Utils;

namespace MIM
{
    class PowerShellInstance
    {
        private string MA_NAME { get; set; }
        public Config Config { get; set; }
        private NLog.Logger logger { get; set; }
        private bool IsDebugEnabled { get; set; }
        private PowerShell PowerShell { get; set; }
        private Dictionary<string, FileInfo> ScriptList { get; set; }

        public PowerShellInstance(string LogName, string MA_NAME)
        {
            this.logger = NLog.LogManager.GetLogger(LogName);
            this.MA_NAME = MA_NAME;

            try
            {
                this.Config = new Config(MA_NAME, null);
                this.IsDebugEnabled = this.Config["IsDebugEnabled"] == "" ? false : bool.Parse(this.Config["IsDebugEnabled"]);
                this.PowerShell = null;
                this.ScriptList = new Dictionary<string, FileInfo>();
            }
            catch (Exception e)
            {
                logger.Error("{0}", e.Message);
                logger.Error("{0}", e.Source);
                logger.Error("{0}", e.StackTrace);
                throw e;
            }
        }

        public void Dispose()
        {
            this.MA_NAME = null;
            this.Config = null;
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
                if (this.IsDebugEnabled)
                {
                    this.logger.Debug("Invoke command: {0} Parameters: {1}", Command, (Parameters == null ? "" : string.Join(",", Parameters.Keys)));
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

            if (this.IsDebugEnabled)
                this.logger.Debug("Invoke command: {0} in {1}ms", Command, (DateTime.Now - StartTime).TotalMilliseconds);

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
                    this.logger.Info("LastWriteTime change on: {0}", scriptfile);
                    if (this.IsDebugEnabled)
                    {
                        this.logger.Debug("{0} : {1}", this.ScriptList[scriptfile].LastWriteTime, newFilinfo.LastWriteTime);
                    }
                    Reload = true;
                    break;
                }
            }

            if (Reload)
            {
                DateTime StartTime = DateTime.Now;

                if (this.IsDebugEnabled)
                    this.logger.Debug("Reload script");

                //Copy and clear list
                string[] ScriptList = new string[this.ScriptList.Keys.Count];
                this.ScriptList.Keys.CopyTo(ScriptList, 0);
                this.ScriptList.Clear();

                DateTime StartPSVariableTime = DateTime.Now;
                //Get PowerShellInstance PSVariable
                Dictionary<string, PSVariable> PSVariableList = new Dictionary<string, PSVariable>();
                Collection<PSObject> list = this.InvokeCommand("Get-Variable", null);
                foreach(PSObject item in list){
                    if (item != null && item.BaseObject != null)
                    {
                        if (item.BaseObject.GetType() == typeof(PSVariable))
                        {
                            PSVariable temp = (PSVariable)item.BaseObject;
                            PSVariableList.Add(temp.Name, temp);

                            if (this.IsDebugEnabled)
                                this.logger.Debug("Save variabel '{0}'", temp.Name);

                        }
                    }
                }

                if (this.IsDebugEnabled)
                    this.logger.Debug("Save variabels in {0}ms", (DateTime.Now - StartPSVariableTime).TotalMilliseconds);


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
                if (this.IsDebugEnabled)
                    logger.Debug("Set PSVariable(s): {0} in {1}ms", string.Join(",", PSVariableList.Keys), (DateTime.Now - StartPSVariableTime).TotalMilliseconds);

                if (this.IsDebugEnabled)
                    this.logger.Debug("Reload script in {0}ms", (DateTime.Now - StartTime).TotalMilliseconds);
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
                if (this.IsDebugEnabled)
                    this.logger.Debug("Create PowerShellInstance");

                this.PowerShell = PowerShell.Create();
                this.PowerShell.Streams.Error.DataAdded += this.Error_DataAdded;
            }

            if (this.IsDebugEnabled)
                this.logger.Debug("Load Script(s)");

            foreach (string scriptfile in ScriptList)
            {
                if (System.IO.File.Exists(scriptfile))
                {
                    
                    if (!this.ScriptList.ContainsKey(scriptfile))
                    {
                        DateTime StartLoadTime = DateTime.Now;
                        if (this.IsDebugEnabled)
                            this.logger.Debug("Load: {0}", scriptfile);

                        //load scriptfile
                        this.PowerShell.AddScript(string.Format(". '{0}'", scriptfile));
                        this.PowerShell.Invoke();
                        this.PowerShell.Commands.Clear();

                        //Add fileinfo to list
                        System.IO.FileInfo fileinfo = new System.IO.FileInfo(scriptfile);
                        this.ScriptList.Add(scriptfile, fileinfo);

                        if (this.IsDebugEnabled)
                            this.logger.Debug("Loaded {0} in : {1}ms", scriptfile,(DateTime.Now-StartLoadTime).TotalMilliseconds);
                    }

                }
                else
                {
                    this.logger.Error("{0} dont exist", scriptfile);
                }
            }



            if (RunInitialize)
            {
                DateTime StartInitializeTime = DateTime.Now;

                if (this.IsDebugEnabled)
                    this.logger.Debug("Initialize script");
                InvokeCommand("Initialize", new Dictionary<string, object>() { { "logger", this.logger }, { "MAName", this.MA_NAME }, { "Config", this.Config } });

                if (this.IsDebugEnabled)
                    this.logger.Debug("Initialize script in {0}ms", (DateTime.Now - StartInitializeTime).TotalMilliseconds);
            }

            if (this.IsDebugEnabled)
                this.logger.Debug("Create PowerShellInstance in {0}ms", (DateTime.Now - StartTime).TotalMilliseconds);
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
        private string MA_NAME;
        private bool IsDebugEnabled = false;
        private string loggerFullName;
        private Config Config;
        private Dictionary<string, PowerShellInstance> PowerShellInstance = new Dictionary<string, PowerShellInstance>();
        private string[] IMVSynchronizationList;

        public Powershell()
        {
            //MA name & Config
            loggerFullName = typeof(Powershell).FullName;
            MA_NAME = loggerFullName.Substring(loggerFullName.LastIndexOf('.') + 1);
            Config = new Config(MA_NAME, null);
            //Log
            
            logger = NLog.LogManager.GetLogger(loggerFullName);
            NLog.LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration(Config.LoggingConfiguration);
            //Debug?
            //IsDebugEnabled = logger.IsDebugEnabled;
            logger.Warn("dafsdf");
            logger.Warn(loggerFullName);
            IsDebugEnabled = bool.Parse(Config["IsDebugEnabled"]);
            
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
                    if (IsDebugEnabled)
                        logger.Debug("Object type : {0}", obj.BaseObject.GetType().Name);

                    if (obj.BaseObject.GetType() == typeof(T))
                    {
                        returnObject = (T)obj.BaseObject;
                        break;
                    }
                }
                else if (IsDebugEnabled)
                    logger.Debug("Object is null!");
            }
            return returnObject;
        }

        #region IMASynchronization

        void IMASynchronization.Initialize()
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Initialize {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
            }
            if (IsDebugEnabled)
                logger.Debug("IMASynchronization path: {0}", CurentPowerShellInstance.Config["IMASynchronization"]);

            CurentPowerShellInstance.LoadScriptList(CurentPowerShellInstance.Config["IMASynchronization"], true);
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.Initialize", null);

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.Initialize {0}", curent_MA_NAME);
        }

        void IMASynchronization.Terminate()
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Terminate {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.Terminate", null);

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.Terminate {0}", curent_MA_NAME);
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.ShouldProjectToMV {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }

            Collection<PSObject> List = CurentPowerShellInstance.InvokeCommand("IMASynchronization.ShouldProjectToMV", new Dictionary<string, object>() { { "CSEntry", csentry } });
            bool result = GetFirstObjectOf<bool>(List);
            MVObjectType = GetFirstObjectOf<string>(List);


            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ShouldProjectToMV {0}", curent_MA_NAME);

            return result;
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Deprovision {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }

            DeprovisionAction result = GetFirstObjectOf<DeprovisionAction>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.Deprovision", new Dictionary<string, object>() { { "CSEntry", csentry } }));

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Deprovision {0}", curent_MA_NAME);

            return result;
        }

        bool IMASynchronization.FilterForDisconnection(CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.FilterForDisconnection {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            bool result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.FilterForDisconnection", new Dictionary<string, object>() { { "CSEntry", csentry } }));

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.FilterForDisconnection {0}", curent_MA_NAME);

            return result;
        }

        void IMASynchronization.MapAttributesForJoin(string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
            {
                logger.Debug("Start IMASynchronization.MapAttributesForJoin {0}", curent_MA_NAME);
                logger.Debug("FlowRuleName: {0} ", FlowRuleName);
            }

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMASynchronization.MapAttributesForJoin", new Dictionary<string, object>() { 
            { "FlowRuleName", FlowRuleName } ,
            { "CSEntry", csentry } ,
            { "ValueCollection", values } 
            });


            if (IsDebugEnabled)
            {
                logger.Debug("ValueCollection: {0}", string.Join(",", values.ToStringArray()));
                logger.Debug("Done IMASynchronization.MapAttributesForJoin {0}", curent_MA_NAME);
            }
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.ResolveJoinSearch {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            Collection<PSObject> List = CurentPowerShellInstance.InvokeCommand("IMASynchronization.ResolveJoinSearch", new Dictionary<string, object>() { 
            { "joinCriteriaName", joinCriteriaName } ,
            { "CSEntry", csentry } ,
            { "rgmventry", rgmventry } 
            });

            bool result = GetFirstObjectOf<bool>(List);
            imventry = GetFirstObjectOf<int>(List);

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ResolveJoinSearch {0}", curent_MA_NAME);

            return result;
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForImport {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
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


            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForImport {0}", curent_MA_NAME);
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForExport {0}", curent_MA_NAME);


            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            bool result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMASynchronization.MapAttributesForExport", new Dictionary<string, object>() { 
            { "FlowRuleName", FlowRuleName } ,
            { "CSEntry", csentry } ,
            { "MVEntry", mventry } 
            }));

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForExport {0}", curent_MA_NAME);
        }

        #endregion

        #region IMVSynchronization

        void IMVSynchronization.Initialize()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Initialize");

            IMVSynchronizationList = Config["IMVSynchronization-MA-List"].Replace("\t", "").Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (IsDebugEnabled)
                    logger.Debug("Start IMVSynchronization.Initialize {0}", curent_MA_NAME);

                PowerShellInstance CurentPowerShellInstance;
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    CurentPowerShellInstance = new PowerShellInstance($"{loggerFullName}.{curent_MA_NAME}", curent_MA_NAME);
                    PowerShellInstance.Add(curent_MA_NAME, CurentPowerShellInstance);
                }

                CurentPowerShellInstance.LoadScriptList(CurentPowerShellInstance.Config["IMVSynchronization"],true);
                CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Initialize", null);
            }

            if (IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Initialize");
        }

        void IMVSynchronization.Terminate()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Terminate");

            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                    throw (new Exception("PowerShell instance is null"));
                }

                CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Terminate", null);

            }

            if (IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Terminate");
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            DateTime StartTime = DateTime.Now;

            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Provision");

            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                    throw (new Exception("PowerShell instance is null"));
                }
                DateTime StartInvokeTime = DateTime.Now;
                if (IsDebugEnabled)
                    logger.Debug("Start IMVSynchronization.Provision on {0} ", curent_MA_NAME);
                CurentPowerShellInstance.InvokeCommand("IMVSynchronization.Provision", new Dictionary<string, object>() { { "MVEntry", mventry } });
                if (IsDebugEnabled)
                    logger.Debug("Start IMVSynchronization.Provision on {0} in {1}ms", curent_MA_NAME, (DateTime.Now - StartInvokeTime).TotalMilliseconds);

            }

            if (IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Provision in {0}ms", (DateTime.Now - StartTime).TotalMilliseconds);
        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.ShouldDeleteFromMV");

            bool result = false;
            PowerShellInstance CurentPowerShellInstance;
            foreach (string curent_MA_NAME in IMVSynchronizationList)
            {
                if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
                {
                    logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                    throw (new Exception("PowerShell instance is null"));
                }
                result = GetFirstObjectOf<bool>(CurentPowerShellInstance.InvokeCommand("IMVSynchronization.ShouldDeleteFromMV", new Dictionary<string, object>() { 
                    { "CSEntry", csentry } ,
                    { "MVEntry", mventry }
                    }));

            }

            if (IsDebugEnabled)
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
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetParameters.GetConfigParameters");

            System.Collections.Generic.List<ConfigParameterDefinition> list;
            string IMAExtensible2Script = configParameters.Contains("IMAExtensible2GetParameters") ? configParameters["IMAExtensible2GetParameters"].Value : Config["IMAExtensible2GetParameters"];

            if (IsDebugEnabled)
                logger.Debug("Load: {0}", IMAExtensible2Script);

            PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}.{MA_NAME}", MA_NAME);
            CurrentPowerShellInstance.LoadScriptList(IMAExtensible2Script,true);
            list = GetFirstObjectOf<System.Collections.Generic.List<ConfigParameterDefinition>>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetParameters.GetConfigParameters", new Dictionary<string, object>() { 
                    { "ConfigParameters", configParameters } ,
                    { "ConfigParameterPage", page }
                    }));

            CurrentPowerShellInstance.Dispose();
            CurrentPowerShellInstance = null;

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.GetConfigParameters");

            return list;
        }

        MACapabilities IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");

            MACapabilities result;
            string IMAExtensible2GetCapabilitiesEx = configParameters["IMAExtensible2GetCapabilitiesEx"].Value;
            if (IMAExtensible2GetCapabilitiesEx.Length > 0)
            {
                if (IsDebugEnabled)
                    logger.Debug("Load: {0}", IMAExtensible2GetCapabilitiesEx);

                PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}.{MA_NAME}", MA_NAME);
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
                    logger.Error("{0}", e.Message);
                    logger.Error("{0}", e.Source);
                    logger.Error("{0}", e.StackTrace);
                    throw new ExtensibleExtensionException(e.Message);
                }
            }

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");

            return result;
        }

        Schema IMAExtensible2GetSchema.GetSchema(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetSchema.GetSchema");

            Schema result;
            if (IsDebugEnabled)
                logger.Debug("Load: {0}", configParameters["IMAExtensible2GetSchema"].Value);

            PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}.{MA_NAME}", MA_NAME);
            CurrentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2GetSchema"].Value,true);
            result = GetFirstObjectOf<Schema>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetSchema.GetSchema", new Dictionary<string, object>() { { "ConfigParameters", configParameters } }));

            CurrentPowerShellInstance.Dispose();
            CurrentPowerShellInstance = null;

            if (IsDebugEnabled)
            {
                //log schema
                foreach (SchemaType type in result.Types)
                {
                    logger.Debug("Schema type: {0}", type.Name);
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug("Schema attibute: {0}", attibute.Name);
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug("Schema AnchorAttributes: {0}", attibute.Name);
                    }
                }

                logger.Debug("Done IMAExtensible2GetSchema.GetSchema");
            }

            return result;
        }

        ParameterValidationResult IMAExtensible2GetParameters.ValidateConfigParameters(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetParameters.ValidateConfigParameters");

            ParameterValidationResult result;

            if (IsDebugEnabled)
                logger.Debug("Load: {0}", configParameters["IMAExtensible2GetParameters"].Value);

            PowerShellInstance CurrentPowerShellInstance = new MIM.PowerShellInstance($"{loggerFullName}.{MA_NAME}", MA_NAME);
            CurrentPowerShellInstance.LoadScriptList(configParameters["IMAExtensible2GetParameters"].Value,true);
            result = GetFirstObjectOf<ParameterValidationResult>(CurrentPowerShellInstance.InvokeCommand("IMAExtensible2GetParameters.ValidateConfigParameters", new Dictionary<string, object>() { 
            { "ConfigParameters", configParameters } ,
            { "ConfigParameterPage", page } 
            }));

            CurrentPowerShellInstance.Dispose();
            CurrentPowerShellInstance = null;

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.ValidateConfigParameters");

            return result;
        }

        #endregion

        #region import

        OpenImportConnectionResults IMAExtensible2CallImport.OpenImportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenImportConnectionRunStep openImportRunStep)
        {

            OpenImportConnectionResults result = null;

            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);



            if (IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallImport.OpenImportConnection {0}", curent_MA_NAME);
                logger.Debug("openImportRunStep: Type: {0} PageSize: {1} CustomData: {2}", openImportRunStep.ImportType, openImportRunStep.PageSize, openImportRunStep.CustomData);

                //log schema
                foreach (SchemaType type in types.Types)
                {
                    logger.Debug("Schema type: {0}", type.Name);
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug("Schema attibute: {0}", attibute.Name);
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug("Schema AnchorAttributes: {0}", attibute.Name);
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

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.OpenImportConnection");

            return result;
        }

        GetImportEntriesResults IMAExtensible2CallImport.GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallImport.GetImportEntries {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            GetImportEntriesResults result = GetFirstObjectOf<GetImportEntriesResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallImport.GetImportEntries", new Dictionary<string, object>() { 
            { "GetImportEntriesRunStep", importRunStep }
            }));

            if (IsDebugEnabled)
            {
                logger.Debug("MoreToImport? {0}", result.MoreToImport);
                logger.Debug("Done IMAExtensible2CallImport.GetImportEntries {0}", curent_MA_NAME);
            }

            return result;
        }

        CloseImportConnectionResults IMAExtensible2CallImport.CloseImportConnection(CloseImportConnectionRunStep importRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallImport.CloseImportConnection {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CloseImportConnectionResults result = GetFirstObjectOf<CloseImportConnectionResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallImport.CloseImportConnection", new Dictionary<string, object>() { 
            { "CloseImportConnectionRunStep", importRunStep }
            }));

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.CloseImportConnection {0}", curent_MA_NAME);

            return result;
        }
        #endregion

        #region export

        void IMAExtensible2CallExport.OpenExportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenExportConnectionRunStep exportRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallExport.OpenExportConnection {0}", curent_MA_NAME);
                logger.Debug("OpenExportConnection: Type: {0} BatchSize: {1} StepPartition: {2}", exportRunStep.ExportType, exportRunStep.BatchSize, exportRunStep.StepPartition.Name);

                //log schema
                foreach (SchemaType type in types.Types)
                {
                    logger.Debug("Schema type: {0}", type.Name);
                    foreach (SchemaAttribute attibute in type.Attributes)
                    {
                        logger.Debug("Schema attibute: {0}", attibute.Name);
                    }
                    foreach (SchemaAttribute attibute in type.AnchorAttributes)
                    {
                        logger.Debug("Schema AnchorAttributes: {0}", attibute.Name);
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

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.OpenExportConnection {0}", curent_MA_NAME);
        }

        PutExportEntriesResults IMAExtensible2CallExport.PutExportEntries(System.Collections.Generic.IList<CSEntryChange> csentries)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallExport.PutExportEntries {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            
            PutExportEntriesResults result = GetFirstObjectOf<PutExportEntriesResults>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallExport.PutExportEntries", new Dictionary<string, object>() { 
            { "CSEntryChanges", csentries }
            }));

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.PutExportEntries {0}", curent_MA_NAME);

            return result;
        }

        void IMAExtensible2CallExport.CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallExport.CloseExportConnection {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2CallExport.CloseExportConnection", new Dictionary<string, object>() { 
            { "CloseExportConnectionRunStep", exportRunStep }
            });

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.CloseExportConnection {0}", curent_MA_NAME);
        }
        #endregion

        #region password
        void IMAExtensible2Password.OpenPasswordConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Partition partition)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.OpenPasswordConnection {0}", curent_MA_NAME);

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

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.OpenPasswordConnection {0}", curent_MA_NAME);
        }

        ConnectionSecurityLevel IMAExtensible2Password.GetConnectionSecurityLevel()
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.GetConnectionSecurityLevel {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            ConnectionSecurityLevel result = GetFirstObjectOf<ConnectionSecurityLevel>(CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.GetConnectionSecurityLevel", null));

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.GetConnectionSecurityLevel {0}", curent_MA_NAME);

            return result;
        }

        void IMAExtensible2Password.ClosePasswordConnection()
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.ClosePasswordConnection {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.ClosePasswordConnection", null);

            /*CurentPowerShellInstance.Dispose();
            CurentPowerShellInstance = null;
            PowerShellInstance.Clear();
            PowerShellInstance = null;*/

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.ClosePasswordConnection {0}", curent_MA_NAME);
        }

        void IMAExtensible2Password.ChangePassword(CSEntry csentry, System.Security.SecureString oldPassword, System.Security.SecureString newPassword)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.ChangePassword {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.ChangePassword", new Dictionary<string, object>() { 
            { "CSEntry", csentry },
            { "oldPassword", oldPassword },
            { "newPassword", newPassword },
            });

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.ChangePassword {0}", curent_MA_NAME);
        }

        void IMAExtensible2Password.SetPassword(CSEntry csentry, System.Security.SecureString newPassword, PasswordOptions options)
        {
            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);

            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.SetPassword {0}", curent_MA_NAME);

            PowerShellInstance CurentPowerShellInstance;
            if (!PowerShellInstance.TryGetValue(curent_MA_NAME, out CurentPowerShellInstance))
            {
                logger.Error("ERROR geting powershell in for {0}", curent_MA_NAME);
                throw (new Exception("PowerShell instance is null"));
            }
            CurentPowerShellInstance.InvokeCommand("IMAExtensible2Password.SetPassword", new Dictionary<string, object>() { 
            { "CSEntry", csentry },
            { "newPassword", newPassword },
            { "PasswordOptions", options },
            });

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.SetPassword {0}", curent_MA_NAME);
        }
        #endregion
    }
}
