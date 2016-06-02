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


using LD.IdentityManagement.Utils;

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