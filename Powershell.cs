/* 2015-02-09 Anders Åsén Landstinget Dalarna
 * 2015-09-17 Add support for run ps on IMASynchronization, IMVSynchronization, IMAExtensible2
 * 2016-02-17 Add support for ReloadPowerShellScript (LastWriteTime)
 *            Add support for DeclineMappingException
 * 2016-05-19 Fix minor error
 * 2016-05-19 Add Config sampel
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.ObjectModel;
using System.Management.Automation;
using Microsoft.MetadirectoryServices;
using System.Xml;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.InteropServices;

namespace LD.IdentityManagement.Utils
{
    namespace Crypto
    {
        public static class ProtectedData
        {
            public static string EncryptString(string input, DataProtectionScope scope, byte[] entropy)
            {
                return Convert.ToBase64String(System.Security.Cryptography.ProtectedData.Protect(System.Text.Encoding.Unicode.GetBytes(input), entropy, scope));
            }

            public static string DecryptString(string encryptedData, DataProtectionScope scope, byte[] entropy)
            {
                return System.Text.Encoding.Unicode.GetString(System.Security.Cryptography.ProtectedData.Unprotect(Convert.FromBase64String(encryptedData), entropy, scope));
            }
        }
    }

    public class Config : ConfigXML
    {
        public Config(string Name, string ConifgPath)
            : base(Name, ConifgPath)
        {

        }
    }

    /*public class ConfigJSON
    {
        private string ConfigName;
        public string LoggingConfiguration;
        Dictionary<string, Object> CurrentConfig = null;

        private string ConfigFilePath = @"C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\conf\LD.IdentityManagement.Agents.Config.json";

        private void Save(Dictionary<string, Object> NewConfig)
        {
            fastJSON.JSONParameters par = new fastJSON.JSONParameters();
            File.WriteAllText(ConfigFilePath, fastJSON.JSON.ToNiceJSON(NewConfig, par), Encoding.UTF8);
        }

        public ConfigJSON()
        { }

        public ConfigJSON(string Name, string ConifgPath)
        {
            if (ConifgPath != null)
                ConfigFilePath = ConifgPath;

            if (File.Exists(ConfigFilePath))
            {
                //object temp = fastJSON.JSON.ToObject(File.ReadAllText(ConfigFilePath));
                Dictionary<string, Object> Confi = fastJSON.JSON.ToObject<Dictionary<string, Object>>(File.ReadAllText(ConfigFilePath, Encoding.UTF8));
                bool dirty = false;

                LoggingConfiguration = ((List<object>)((Dictionary<string, Object>)Confi["LoggingConfiguration"])["File"])[0].ToString();
                ConfigName = Name;
                object temp;
                if (Confi.TryGetValue(Name, out temp))
                {
                    CurrentConfig = (Dictionary<string, Object>)temp;

                    foreach (string Key in CurrentConfig.Keys)
                    {
                        object[] Value = ((List<object>)CurrentConfig[Key]).ToArray();
                        if (Value.Length > 1 && (Value[1].ToString() == "" || Value[1] == null))
                        {
                            Value[0] = LD.IdentityManagement.Utils.Crypto.ProtectedData.EncryptString(Value[0].ToString(), System.Security.Cryptography.DataProtectionScope.CurrentUser, null);
                            Value[1] = "ProtectedData";
                            CurrentConfig[Key] = Value;
                            dirty = true;
                        }
                    }

                    if (dirty)
                    {
                        Confi[Name] = CurrentConfig;
                        Save(Confi);
                    }
                }
            }
        }

        public string Get(String Name)
        {
            String returnValue = "";

            if (CurrentConfig != null)
            {
                Object temp;
                if (CurrentConfig.TryGetValue(Name, out temp))
                {
                    object[] Value = ((List<object>)temp).ToArray();
                    if (Value != null)
                    {
                        if (Value.Length > 1 && Value[1].ToString() == "ProtectedData")
                        {
                            returnValue = LD.IdentityManagement.Utils.Crypto.ProtectedData.DecryptString(Value[0].ToString(), System.Security.Cryptography.DataProtectionScope.CurrentUser, null);
                        }
                        else
                            returnValue = Value[0].ToString();
                    }
                }
            }
            return returnValue;
        }

        public object[] GetRwa(String Name)
        {
            object[] returnValue = null;

            if (CurrentConfig != null)
            {
                returnValue = ((List<object>)CurrentConfig[Name]).ToArray();
            }
            return returnValue;
        }

        public bool Set(String Name, string[] NewValue)
        {
            bool returnValue = false;

            if (CurrentConfig != null && NewValue != null)
            {
                if (File.Exists(ConfigFilePath))
                {
                    Dictionary<string, Object> Confi = fastJSON.JSON.ToObject<Dictionary<string, Object>>(File.ReadAllText(ConfigFilePath, Encoding.UTF8));

                    if (NewValue.Length > 1 && (NewValue[1].ToString() == "" || NewValue[1] == null))
                    {
                        NewValue[0] = LD.IdentityManagement.Utils.Crypto.ProtectedData.EncryptString(NewValue[0], System.Security.Cryptography.DataProtectionScope.CurrentUser, null);
                        NewValue[1] = "ProtectedData";
                    }

                    CurrentConfig[Name] = NewValue;

                    Confi[ConfigName] = CurrentConfig;
                    Save(Confi);
                    returnValue = true;
                }
            }
            return returnValue;
        }

        public bool SetRwa(String Name, string[] NewValue)
        {
            bool returnValue = false;

            if (CurrentConfig != null && NewValue != null)
            {
                if (File.Exists(ConfigFilePath))
                {
                    Dictionary<string, Object> Confi = fastJSON.JSON.ToObject<Dictionary<string, Object>>(File.ReadAllText(ConfigFilePath, Encoding.UTF8));

                    CurrentConfig[Name] = NewValue;

                    Confi[ConfigName] = CurrentConfig;
                    Save(Confi);
                    returnValue = true;
                }
            }
            return returnValue;
        }
    }*/

    public class ConfigXML
    {
        private XmlNode CurrentNode;
        private XmlNode ExtraNode;
        public string LoggingConfiguration;
        private DateTime FileTimeStamp;
        private string CurrentNodeName;

        private string ConfigFilePath = @"C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\conf\LD.IdentityManagement.Agents.Config.xml";


        public ConfigXML()
        { }

        public string this[string name]
        {
            get
            {
                string returnValue = "";
                if (File.Exists(ConfigFilePath))
                {
                    //reload XML
                    if (File.GetLastWriteTime(ConfigFilePath) > FileTimeStamp)
                        LoadXML();

                    if (CurrentNode != null && CurrentNode[name] != null)
                    {
                        returnValue = getInnerString(CurrentNode[name]);
                    }
                    else if (ExtraNode != null && ExtraNode[name] != null)
                    {
                        returnValue = getInnerString(ExtraNode[name]);
                    }
                }
                return returnValue;
            }
            set
            {
                if (value != null && File.Exists(ConfigFilePath))
                {
                    //reload XML
                    if (File.GetLastWriteTime(ConfigFilePath) > FileTimeStamp)
                        LoadXML();

                    if (CurrentNode[name] == null)
                    {
                        CurrentNode.AppendChild(CurrentNode.OwnerDocument.CreateElement(name));
                    }
                    if (CurrentNode.Attributes["ProtectedType"] != null && CurrentNode.Attributes["ProtectedType"].Value == "ProtectedData")
                    {
                        CurrentNode[name].InnerText = LD.IdentityManagement.Utils.Crypto.ProtectedData.EncryptString(value, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

                        if (CurrentNode[name].Attributes["Protected"] == null)
                            CurrentNode[name].Attributes.Append(CurrentNode.OwnerDocument.CreateAttribute("Protected"));

                        if (CurrentNode[name].Attributes["ProtectedType"] == null)
                            CurrentNode[name].Attributes.Append(CurrentNode.OwnerDocument.CreateAttribute("ProtectedType"));

                        CurrentNode[name].Attributes["Protected"].Value = "yes";
                        CurrentNode[name].Attributes["ProtectedType"].Value = "ProtectedData";
                    }
                    else
                    {
                        CurrentNode[name].InnerText = value;
                        if (CurrentNode[name].Attributes["Protected"] != null)
                            CurrentNode[name].RemoveAttribute("Protected");
                        if (CurrentNode[name].Attributes["ProtectedType"] == null)
                            CurrentNode[name].RemoveAttribute("ProtectedType");
                    }

                    CurrentNode.OwnerDocument.Save(ConfigFilePath);
                }
            }
        }


        public ConfigXML(string NodeName, string ConifgPath)
        {

            if (ConifgPath != null)
                ConfigFilePath = ConifgPath;

            CurrentNodeName = NodeName;
            LoadXML();
        }

        private void LoadXML()
        {
            if (File.Exists(ConfigFilePath))
            {
                FileTimeStamp = File.GetLastWriteTime(ConfigFilePath);

                XmlDocument XmlConfig = new XmlDocument();
                XmlConfig.Load(ConfigFilePath);

                LoggingConfiguration = XmlConfig["LD.IdentityManagement"]["LoggingConfiguration"].InnerText;

                bool dirty = false;
                if (XmlConfig["LD.IdentityManagement"]["FIM"] != null)
                {
                    ExtraNode = XmlConfig["LD.IdentityManagement"]["FIM"];
                    dirty |= SetProtectedNode(ExtraNode);
                }


                if (XmlConfig["LD.IdentityManagement"]["Agents"][CurrentNodeName] != null)
                {
                    CurrentNode = XmlConfig["LD.IdentityManagement"]["Agents"][CurrentNodeName];
                    dirty |= SetProtectedNode(CurrentNode);
                }
                //Dirty save config
                if (dirty)
                    XmlConfig.Save(ConfigFilePath);
            }
        }

        private bool SetProtectedNode(XmlNode Top)
        {
            bool dirty = false;
            foreach (XmlNode Node in Top.ChildNodes)
            {
                if (Node.Attributes.Count > 0 && Node.Attributes["Protected"] != null && Node.Attributes["Protected"].Value.ToLower() == "yes" && Node.Attributes["ProtectedType"] == null)
                {
                    Node.InnerText = LD.IdentityManagement.Utils.Crypto.ProtectedData.EncryptString(Node.InnerText, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

                    Node.Attributes.Append(Node.OwnerDocument.CreateAttribute("ProtectedType"));

                    Node.Attributes["ProtectedType"].Value = "ProtectedData";
                    dirty = true;
                }
            }
            return dirty;
        }

        private string getInnerString(XmlNode Node)
        {
            string returnValue = "";
            if (Node.Attributes.Count > 0 && Node.Attributes["Protected"] != null && Node.Attributes["Protected"].Value.ToLower() == "yes")
            {
                if (Node.Attributes["ProtectedType"] != null && Node.Attributes["ProtectedType"].Value == "ProtectedData")
                {
                    returnValue = LD.IdentityManagement.Utils.Crypto.ProtectedData.DecryptString(Node.InnerText, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);
                }
            }
            else
                returnValue = Node.InnerText;

            return returnValue;
        }

        public string Get(String Name)
        {
            //reload XML
            if (File.GetLastWriteTime(ConfigFilePath) > FileTimeStamp)
                LoadXML();

            String returnValue = "";
            if (CurrentNode != null && CurrentNode[Name] != null)
            {
                returnValue = getInnerString(CurrentNode[Name]);
            }
            else if (ExtraNode != null && ExtraNode[Name] != null)
            {
                returnValue = getInnerString(ExtraNode[Name]);
            }
            return returnValue;
        }

        public bool Set(String Name, string NewValue, bool ProtectedData)
        {
            bool returnValue = false;

            if (CurrentNode != null && NewValue != null && File.Exists(ConfigFilePath))
            {
                if (CurrentNode[Name] == null)
                {
                    CurrentNode.AppendChild(CurrentNode.OwnerDocument.CreateElement(Name));
                }
                if (ProtectedData)
                {
                    CurrentNode[Name].InnerText = LD.IdentityManagement.Utils.Crypto.ProtectedData.EncryptString(NewValue, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

                    if (CurrentNode[Name].Attributes["Protected"] == null)
                        CurrentNode[Name].Attributes.Append(CurrentNode.OwnerDocument.CreateAttribute("Protected"));

                    if (CurrentNode[Name].Attributes["ProtectedType"] == null)
                        CurrentNode[Name].Attributes.Append(CurrentNode.OwnerDocument.CreateAttribute("ProtectedType"));

                    CurrentNode[Name].Attributes["Protected"].Value = "yes";
                    CurrentNode[Name].Attributes["ProtectedType"].Value = "ProtectedData";
                }
                else
                {
                    CurrentNode[Name].InnerText = NewValue;
                    if (CurrentNode[Name].Attributes["Protected"] != null)
                        CurrentNode[Name].RemoveAttribute("Protected");
                    if (CurrentNode[Name].Attributes["ProtectedType"] == null)
                        CurrentNode[Name].RemoveAttribute("ProtectedType");
                }

                CurrentNode.OwnerDocument.Save(ConfigFilePath);
                returnValue = true;
            }
            return returnValue;
        }
    }
}


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
        private Dictionary<string, PowerShellConfig> PowerShellInstanceFiles = new Dictionary<string, PowerShellConfig>();
        private PowerShell[] IMVSynchronizationPowerShellInstances = null;

        class PowerShellConfig
        {
            public PowerShellConfig(DateTime LastWriteTime, string MA_NAME, LD.IdentityManagement.Utils.Config Config, NLog.Logger logger)
            {
                this.LastWriteTime = LastWriteTime;
                this.MA_NAME = MA_NAME;
                this.Config = Config;
                this.logger = logger;
            }

            public DateTime LastWriteTime { get; set; }
            public string MA_NAME { get; set; }
            public LD.IdentityManagement.Utils.Config Config { get; set; }
            public NLog.Logger logger { get; set; }

        }

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
            //IsDebugEnabled = logger.IsDebugEnabled;
            IsDebugEnabled = bool.Parse(Config["IsDebugEnabled"]);
        }

        private void ReloadPowerShellScript(PowerShell PowerShellInstance)
        {
            string[] keys = new string[PowerShellInstanceFiles.Keys.Count];
            PowerShellInstanceFiles.Keys.CopyTo(keys, 0);
            foreach (string key in keys)
            {
                System.IO.FileInfo newFil = new System.IO.FileInfo(key);
                if (PowerShellInstanceFiles[key].LastWriteTime != newFil.LastWriteTime)
                {
                    //logger.Debug("{0} {1}", OldFil.FullName, newFil.FullName);
                    //logger.Debug("{0} : {1}", OldFil.LastWriteTime, newFil.LastWriteTime);
                    if (IsDebugEnabled)
                        logger.Debug("ReloadPowerShellScript: {0}", key);

                    PowerShellInstance.AddScript(string.Format(". '{0}'", key));

                    PowerShellInstance.Invoke();
                    PowerShellInstanceFiles[key].LastWriteTime = newFil.LastWriteTime;
                    PowerShellInstance.Commands.Clear();

                    string curent_MA_NAME = Microsoft.MetadirectoryServices.Utils.WorkingDirectory.Substring(Microsoft.MetadirectoryServices.Utils.WorkingDirectory.LastIndexOf('\\') + 1);
                    CurentMAConfig = new LD.IdentityManagement.Utils.Config(curent_MA_NAME, null);
                    NLog.Logger Scriptlogger = curent_MA_NAME == MA_NAME ? logger : NLog.LogManager.GetLogger(loggerFullName + "." + curent_MA_NAME);

                    PowerShellInstance.AddCommand("Initialize");
                    PowerShellInstance.AddParameter("logger", PowerShellInstanceFiles[key].logger);
                    PowerShellInstance.AddParameter("MAName", PowerShellInstanceFiles[key].MA_NAME);
                    PowerShellInstance.AddParameter("Config", PowerShellInstanceFiles[key].Config);
                    PowerShellInstance.Invoke();
                    PowerShellInstance.Commands.Clear();
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

            //NLog logger
            NLog.Logger Scriptlogger = CurrentMA_NAME == MA_NAME ? logger : NLog.LogManager.GetLogger(loggerFullName + "." + CurrentMA_NAME);

            //Load scripts files from list
            foreach (string script in ScriptList)
            {
                if (System.IO.File.Exists(script))
                {
                    System.IO.FileInfo fileinfo = new System.IO.FileInfo(script);
                    if (IsDebugEnabled)
                    {
                        logger.Debug("PowerShellInstanceFiles:{0} exist in list:{1}", script, PowerShellInstanceFiles.ContainsKey(script));
                        logger.Debug("PowerShellInstanceFiles list: {0}", string.Join(",", PowerShellInstanceFiles.Keys));
                    }
                    if (!PowerShellInstanceFiles.ContainsKey(script))
                    {
                        PowerShellInstanceFiles.Add(script, new Powershell.PowerShellConfig(fileinfo.LastWriteTime, CurrentMA_NAME, ScriptConfig, Scriptlogger));
                    }
                    else
                    {
                        PowerShellInstanceFiles[script].LastWriteTime = fileinfo.LastWriteTime;
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

            PowerShellCurent.AddCommand("Initialize");
            PowerShellCurent.AddParameter("logger", Scriptlogger);
            PowerShellCurent.AddParameter("MAName", CurrentMA_NAME);
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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

            PowerShellInstance.AddCommand("IMASynchronization.ShouldProjectToMV");
            PowerShellInstance.AddParameter("CSEntry", csentry);
            //PowerShellInstance.AddParameter("MVObjectType", MVObjectType);
            Collection<PSObject> List = PowerShellInstance.Invoke();
            bool result = GetFirstObjectOf<bool>(List);
            MVObjectType = GetFirstObjectOf<string>(List);
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ShouldProjectToMV");

            return result;
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.Deprovision");
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            {
                logger.Debug("Start IMASynchronization.MapAttributesForJoin");
                logger.Debug("FlowRuleName: {0} ", FlowRuleName);
            }
            ReloadPowerShellScript(PowerShellInstance);

            PowerShellInstance.AddCommand("IMASynchronization.MapAttributesForJoin");
            PowerShellInstance.AddParameter("FlowRuleName", FlowRuleName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("ValueCollection", values);
            PowerShellInstance.Invoke();
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
            {
                logger.Debug("ValueCollection: {0}", string.Join(",", values.ToStringArray()));
                logger.Debug("Done IMASynchronization.MapAttributesForJoin");
            }
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.ResolveJoinSearch");
            ReloadPowerShellScript(PowerShellInstance);

            PowerShellInstance.AddCommand("IMASynchronization.ResolveJoinSearch");
            PowerShellInstance.AddParameter("joinCriteriaName", joinCriteriaName);
            PowerShellInstance.AddParameter("CSEntry", csentry);
            PowerShellInstance.AddParameter("rgmventry", rgmventry);
            //imventry = -1;
            //PowerShellInstance.AddParameter("imventry", imventry);
            PowerShellInstance.AddParameter("MVObjectType", MVObjectType);
            Collection<PSObject> List = PowerShellInstance.Invoke();
            bool result = GetFirstObjectOf<bool>(List);
            imventry = GetFirstObjectOf<int>(List);
            //MVObjectType = GetFirstObjectOf<string>(List);
            PowerShellInstance.Commands.Clear();

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.ResolveJoinSearch");

            return result;
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForImport");
            ReloadPowerShellScript(PowerShellInstance);

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
            finally
            {
                PowerShellInstance.Commands.Clear();
            }

            if (IsDebugEnabled)
                logger.Debug("Done IMASynchronization.MapAttributesForImport");
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            if (IsDebugEnabled)
                logger.Debug("Start IMASynchronization.MapAttributesForExport");
            ReloadPowerShellScript(PowerShellInstance);

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
                ReloadPowerShellScript(Instance);
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
                ReloadPowerShellScript(Instance);
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
                ReloadPowerShellScript(Instance);
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
                logger.Debug("Start IMAExtensible2CallImport.GetImportEntries");
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
            ReloadPowerShellScript(PowerShellInstance);

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
