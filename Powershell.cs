/* 2015-02-09 Anders Åsén Landstinget Dalarna
 * 2015-09-17 Add support for run ps on IMASynchronization, IMVSynchronization, IMAExtensible2
 * 2016-02-17 Add support for ReloadPowerShellScript (LastWriteTime)
 *            Add support for DeclineMappingException
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.ObjectModel;
using System.Management.Automation;
using Microsoft.MetadirectoryServices;

namespace LD.IdentityManagement.Agent
{
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
        private LD.IdentityManagement.Utils.Config Config;
        private LD.IdentityManagement.Utils.Config CurentMAConfig;
        private PowerShell PowerShellInstance = null;
        private Dictionary<string, System.IO.FileInfo> PowerShellInstanceFiles = new Dictionary<string, System.IO.FileInfo>();
        private PowerShell[] IMVSynchronizationPowerShellInstances = null;

        public Powershell()
        {
            //MA name & Config
            loggerFullName = typeof(Powershell).FullName;
            MA_NAME = loggerFullName.Substring(loggerFullName.LastIndexOf('.') + 1);
            Config = new LD.IdentityManagement.Utils.Config(MA_NAME, null);
            
            //Log
            logger = NLog.LogManager.GetLogger(loggerFullName);
            NLog.LogManager.Configuration = new NLog.Config.XmlLoggingConfiguration(Config.LoggingConfiguration);
            //Debug?
            IsDebugEnabled = logger.IsDebugEnabled;
        }

        private void ReloadPowerShellScript()
        {
            string[] keys = new string[PowerShellInstanceFiles.Keys.Count];
            PowerShellInstanceFiles.Keys.CopyTo(keys, 0);
            foreach (string key in keys)
            {
                System.IO.FileInfo OldFil = PowerShellInstanceFiles[key];
                System.IO.FileInfo newFil = new System.IO.FileInfo(OldFil.FullName);
                if(OldFil.LastWriteTime != newFil.LastWriteTime)
                {
                    //logger.Debug("{0} {1}", OldFil.FullName, newFil.FullName);
                    //logger.Debug("{0} : {1}", OldFil.LastWriteTime, newFil.LastWriteTime);

                    PowerShellInstance.AddScript(string.Format(". '{0}'", OldFil.FullName));
                    PowerShellInstance.Invoke();
                    PowerShellInstanceFiles[key] = newFil;
                }
            }
        }

        /// <summary>
        /// Initializ PowerShell instance
        /// return true if sucess
        /// </summary>
        /// <param name="ScriptList"></param>
        /// <param name="CurrentMA_NAME"></param>
        /// <returns></returns>
        private PowerShell InitializePS(string[] ScriptList, string CurrentMA_NAME, LD.IdentityManagement.Utils.Config ScriptConfig)
        {
            if (IsDebugEnabled)
            {
                logger.Debug("Start InitializePS");
                logger.Debug("{0} {1} {2}", string.Join(",", ScriptList), CurrentMA_NAME, ScriptConfig.LoggingConfiguration);
            }
            //Create PowerShellInstance
            PowerShell PowerShellCurent = PowerShell.Create();
            PowerShellCurent.Streams.Error.DataAdded += Error_DataAdded;

            //Load scripts files from list
            foreach (string script in ScriptList)
            {
                if (System.IO.File.Exists(script))
                {
                    System.IO.FileInfo fileinfo = new System.IO.FileInfo(script);
                    if (PowerShellInstanceFiles.ContainsKey(script))
                    {
                        PowerShellInstanceFiles.Add(script, fileinfo);
                    }
                    else
                    {
                        PowerShellInstanceFiles[script] = fileinfo;
                    }

                    PowerShellCurent.AddScript(string.Format(". '{0}'", script));
                    PowerShellCurent.Invoke();
                }
                else
                {
                    logger.Error("File dont exist {0}", script);
                    throw new Exception("File dont exist " + script);
                }
            }

            //NLog logger
            NLog.Logger Scriptlogger = CurrentMA_NAME == MA_NAME ? logger : NLog.LogManager.GetLogger(loggerFullName + "." + CurrentMA_NAME);

            PowerShellCurent.AddCommand("Initialize");
            PowerShellCurent.AddParameter("logger", Scriptlogger);
            PowerShellCurent.AddParameter("MA-Name", CurrentMA_NAME);
            PowerShellCurent.AddParameter("Config", ScriptConfig);
            PowerShellCurent.Invoke();
            PowerShellCurent.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done InitializePS");

            return PowerShellCurent;
        }

        /// <summary>
        /// Error logger
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Error_DataAdded(object sender, DataAddedEventArgs e)
        {
            PSDataCollection<ErrorRecord> ErrorRecords = (PSDataCollection<ErrorRecord>)sender;
            ErrorRecord ErrorRecord = ErrorRecords[e.Index];
            if (ErrorRecord != null)
            {
                logger.Error(ErrorRecord.Exception.Message);
                logger.Error(ErrorRecord.Exception.Source);
                logger.Error(ErrorRecord.Exception.StackTrace);
                logger.Error(ErrorRecord.ScriptStackTrace);
                ErrorRecords.RemoveAt(e.Index);
            }
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
                //if (IsDebugEnabled)
                    logger.Debug("Object type : {0}", obj.BaseObject.GetType().Name);

                if (obj.BaseObject.GetType() == typeof(T))
                {
                    returnObject = (T)obj.BaseObject;
                    break;
                }
            }
            return returnObject;
        }

        #region IMASynchronization

        void IMASynchronization.Initialize()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Initialize");

            string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);
            CurentMAConfig = new LD.IdentityManagement.Utils.Config(curent_MA_NAME, null);
            PowerShellInstance = InitializePS(new string[] { CurentMAConfig["IMASynchronization"] }, curent_MA_NAME, CurentMAConfig);

            PowerShellInstance.AddCommand("IMASynchronization.Initialize");
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.Initialize");
        }

        void IMASynchronization.Terminate()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Terminate");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.Terminate");
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.Terminate");
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.ShouldProjectToMV");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.ShouldProjectToMV");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            MVObjectType = "";
            PowerShellInstance.AddParameter("MVObjectType", MVObjectType);
            bool result = GetFirstObjectOf<bool>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ShouldProjectToMV");

            return result;
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Deprovision");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.Deprovision");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            DeprovisionAction result = GetFirstObjectOf<DeprovisionAction>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Deprovision");

            return result;
        }

        bool IMASynchronization.FilterForDisconnection(CSEntry csentry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.FilterForDisconnection");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.FilterForDisconnection");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            bool result = GetFirstObjectOf<bool>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.FilterForDisconnection");

            return result;
        }

        void IMASynchronization.MapAttributesForJoin(string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForJoin");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.MapAttributesForJoin");
            PowerShellInstance.AddParameter("FlowRuleName", FlowRuleName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("ValueCollection", values);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForJoin");
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.ResolveJoinSearch");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.ResolveJoinSearch");
            PowerShellInstance.AddParameter("joinCriteriaName", joinCriteriaName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("rgmventry", rgmventry);
            imventry = -1;
            PowerShellInstance.AddParameter("imventry", imventry);
            PowerShellInstance.AddParameter("MVObjectType", MVObjectType);
            bool result = GetFirstObjectOf<bool>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ResolveJoinSearch");

            return result;
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForImport");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.MapAttributesForImport");
            PowerShellInstance.AddParameter("FlowRuleName", FlowRuleName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("MVEntry", mventry);

            try
            {
                bool result = GetFirstObjectOf<bool>(PowerShellInstance.Invoke());
            }
            catch (Exception e)
            {
                logger.Error("{0}\n{1}\n{2}\n{3}", e.GetType(), e.Message, e.Source, e.StackTrace);

                if (e.Message == "Microsoft.MetadirectoryServices.DeclineMappingException")
                {
                    throw new Microsoft.MetadirectoryServices.DeclineMappingException();
                }
                else
                {
                    throw;
                }
            }
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForImport");
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForExport");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMASynchronization.MapAttributesForExport");
            PowerShellInstance.AddParameter("FlowRuleName", FlowRuleName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("MVEntry", mventry);

            bool result = GetFirstObjectOf<bool>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForExport");
        }

        #endregion

        #region IMVSynchronization

        void IMVSynchronization.Initialize()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Initialize");

            string[] MAlist = Config["IMVSynchronization-MA-List"].Replace("\t", "").Split(new char[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            IMVSynchronizationPowerShellInstances = new PowerShell[MAlist.Length];

            for(int i = 0; i < MAlist.Length; i++)
            {
                CurentMAConfig = new LD.IdentityManagement.Utils.Config(MAlist[i], null);
                IMVSynchronizationPowerShellInstances[i] = InitializePS(new string[] { CurentMAConfig["IMVSynchronization"] }, MAlist[i], CurentMAConfig);
                if (IMVSynchronizationPowerShellInstances[i] != null)
                {
                    IMVSynchronizationPowerShellInstances[i].AddCommand("IMVSynchronization.Initialize");
                    IMVSynchronizationPowerShellInstances[i].Invoke();
                    IMVSynchronizationPowerShellInstances[i].Commands.Clear();
                }
                else
                    throw new Exception("Error loading script file IMVSynchronization : " + MAlist[i]);
            }

            if (IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Initialize");
        }

        void IMVSynchronization.Terminate()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Terminate");

            foreach (PowerShell Instance in IMVSynchronizationPowerShellInstances)
            {
                Instance.AddCommand("IMVSynchronization.Terminate");
                Instance.Invoke();
                Instance.Commands.Clear();
            }

            if (IsDebugEnabled)
                logger.Debug("Done IMVSynchronization.Terminate");
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Provision");

            foreach (PowerShell Instance in IMVSynchronizationPowerShellInstances)
            {
                Instance.AddCommand("IMVSynchronization.Provision");
                Instance.AddParameter("MVEntry", mventry);
                Instance.Invoke();
                Instance.Commands.Clear();
            }

            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.Provision");
        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMVSynchronization.ShouldDeleteFromMV");

            bool result = false;

            foreach (PowerShell Instance in IMVSynchronizationPowerShellInstances)
            {
                Instance.AddCommand("IMVSynchronization.ShouldDeleteFromMV");
                Instance.AddParameter("CSEntry", csentry);
                Instance.AddParameter("MVEntry", mventry);
                if (GetFirstObjectOf<bool>(Instance.Invoke()))
                    return true;
                Instance.Commands.Clear();
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
            PowerShellInstance = InitializePS(new string[] { IMAExtensible2Script }, "IMAExtensible2GetParameters", Config);

            PowerShellInstance.AddCommand("IMAExtensible2GetParameters.GetConfigParameters");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            PowerShellInstance.AddParameter("ConfigParameterPage", page);
            list = GetFirstObjectOf<System.Collections.Generic.List<ConfigParameterDefinition>>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;


            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.GetConfigParameters");

            return list;
        }

        MACapabilities IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");

            MACapabilities result;

            if (configParameters["IMAExtensible2GetCapabilitiesEx"].Value.Length > 0)
            {

                PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2GetCapabilitiesEx"].Value }, "IMAExtensible2GetCapabilitiesEx", Config);

                PowerShellInstance.AddCommand("IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx");
                result = GetFirstObjectOf<MACapabilities>(PowerShellInstance.Invoke());
                PowerShellInstance.Commands.Clear();

                //Close PS instance
                PowerShellInstance.Dispose();
                PowerShellInstance = null;
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
            PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2GetSchema"].Value }, "IMAExtensible2GetSchema", Config);

            PowerShellInstance.AddCommand("IMAExtensible2GetSchema.GetSchema");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            result = GetFirstObjectOf<Schema>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;


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
            PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2GetParameters"].Value }, "IMAExtensible2GetParameters", Config);

            PowerShellInstance.AddCommand("IMAExtensible2GetParameters.ValidateConfigParameters");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            PowerShellInstance.AddParameter("ConfigParameterPage", page);
            result = GetFirstObjectOf<ParameterValidationResult>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;


            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2GetParameters.ValidateConfigParameters");

            return result;
        }

        #endregion

        #region import

        OpenImportConnectionResults IMAExtensible2CallImport.OpenImportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenImportConnectionRunStep openImportRunStep)
        {
            if (IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallImport.OpenImportConnection");
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

            OpenImportConnectionResults result = null;

            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);
            CurentMAConfig = new LD.IdentityManagement.Utils.Config(curent_MA_NAME, null);
            PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2CallImport"].Value }, curent_MA_NAME, CurentMAConfig);

            PowerShellInstance.AddCommand("IMAExtensible2CallImport.OpenImportConnection");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            PowerShellInstance.AddParameter("Schema", types);
            PowerShellInstance.AddParameter("OpenImportConnectionRunStep", openImportRunStep);
            result = GetFirstObjectOf<OpenImportConnectionResults>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.OpenImportConnection");

            return result;
        }

        GetImportEntriesResults IMAExtensible2CallImport.GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            if (IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallImport.GetImportEntries");
            }

            PowerShellInstance.AddCommand("IMAExtensible2CallImport.GetImportEntries");
            PowerShellInstance.AddParameter("GetImportEntriesRunStep", importRunStep);
            GetImportEntriesResults result = GetFirstObjectOf<GetImportEntriesResults>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.GetImportEntries");

            return result;
        }

        CloseImportConnectionResults IMAExtensible2CallImport.CloseImportConnection(CloseImportConnectionRunStep importRunStep)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallImport.CloseImportConnection");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2CallImport.CloseImportConnection");
            PowerShellInstance.AddParameter("CloseImportConnectionRunStep", importRunStep);
            CloseImportConnectionResults result = GetFirstObjectOf<CloseImportConnectionResults>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallImport.CloseImportConnection");

            return result;
        }
        #endregion

        #region export

        void IMAExtensible2CallExport.OpenExportConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenExportConnectionRunStep exportRunStep)
        {
            if (IsDebugEnabled)
            {
                logger.Debug("Start IMAExtensible2CallExport.OpenExportConnection");
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

            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);
            CurentMAConfig = new LD.IdentityManagement.Utils.Config(curent_MA_NAME, null);

            PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2CallExport"].Value }, curent_MA_NAME, CurentMAConfig);

            PowerShellInstance.AddCommand("IMAExtensible2CallExport.OpenExportConnection");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            PowerShellInstance.AddParameter("Schema", types);
            PowerShellInstance.AddParameter("OpenExportConnectionRunStep", exportRunStep);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.OpenExportConnection");
        }

        PutExportEntriesResults IMAExtensible2CallExport.PutExportEntries(System.Collections.Generic.IList<CSEntryChange> csentries)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallExport.PutExportEntries");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2CallExport.PutExportEntries");
            PowerShellInstance.AddParameter("CSEntryChanges", csentries);
            PutExportEntriesResults result = GetFirstObjectOf<PutExportEntriesResults>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.PutExportEntries");

            return result;
        }

        void IMAExtensible2CallExport.CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2CallExport.CloseExportConnection");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2CallExport.CloseExportConnection");
            PowerShellInstance.AddParameter("CloseExportConnectionRunStep", exportRunStep);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2CallExport.CloseExportConnection");
        }
        #endregion

        #region password
        void IMAExtensible2Password.OpenPasswordConnection(System.Collections.ObjectModel.KeyedCollection<string, ConfigParameter> configParameters, Partition partition)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.OpenPasswordConnection");

            string curent_MA_NAME = MAUtils.MAFolder.Substring(MAUtils.MAFolder.LastIndexOf('\\') + 1);
            CurentMAConfig = new LD.IdentityManagement.Utils.Config(curent_MA_NAME, null);
            PowerShellInstance = InitializePS(new string[] { configParameters["IMAExtensible2Password"].Value }, curent_MA_NAME, CurentMAConfig);

            //Exec OpenImport
            PowerShellInstance.AddCommand("IMAExtensible2Password.OpenPasswordConnection");
            PowerShellInstance.AddParameter("ConfigParameters", configParameters);
            PowerShellInstance.AddParameter("Partition", partition);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.OpenPasswordConnection");
        }

        ConnectionSecurityLevel IMAExtensible2Password.GetConnectionSecurityLevel()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.GetConnectionSecurityLevel");

            PowerShellInstance.AddCommand("IMAExtensible2Password.GetConnectionSecurityLevel");
            ConnectionSecurityLevel result = GetFirstObjectOf<ConnectionSecurityLevel>(PowerShellInstance.Invoke());
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.GetConnectionSecurityLevel");

            return result;
        }

        void IMAExtensible2Password.ClosePasswordConnection()
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.ClosePasswordConnection");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2Password.ClosePasswordConnection");
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            //Close PS instance
            PowerShellInstance.Dispose();
            PowerShellInstance = null;

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.ClosePasswordConnection");
        }

        void IMAExtensible2Password.ChangePassword(CSEntry csentry, System.Security.SecureString oldPassword, System.Security.SecureString newPassword)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.ChangePassword");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2Password.ChangePassword");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("oldPassword", oldPassword);
            PowerShellInstance.AddParameter("newPassword", newPassword);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.ChangePassword");
        }

        void IMAExtensible2Password.SetPassword(CSEntry csentry, System.Security.SecureString newPassword, PasswordOptions options)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMAExtensible2Password.SetPassword");
            ReloadPowerShellScript();

            PowerShellInstance.AddCommand("IMAExtensible2Password.SetPassword");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("newPassword", newPassword);
            PowerShellInstance.AddParameter("PasswordOptions", options);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMAExtensible2Password.SetPassword");
        }
        #endregion
    }
}
