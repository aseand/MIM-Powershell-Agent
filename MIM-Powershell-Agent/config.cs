using System;
using System.Xml;
using System.IO;
using System.Security.Cryptography;

namespace Utils
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

    public class ConfigXML
    {
        private XmlNode CurrentNode;
        private XmlNode ExtraNode;
        public string LoggingConfiguration;
        private DateTime FileTimeStamp;
        private string CurrentNodeName;
        private string ConfigFilePath = @"C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\conf\IdentityManagement.Agents.Config.xml";

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
                        CurrentNode[name].InnerText = Utils.Crypto.ProtectedData.EncryptString(value, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

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

                LoggingConfiguration = XmlConfig["IdentityManagement"]["LoggingConfiguration"].InnerText;

                bool dirty = false;
                if (XmlConfig["IdentityManagement"]["MIM"] != null)
                {
                    ExtraNode = XmlConfig["IdentityManagement"]["MIM"];
                    dirty |= SetProtectedNode(ExtraNode);
                }


                if (XmlConfig["IdentityManagement"]["Agents"][CurrentNodeName] != null)
                {
                    CurrentNode = XmlConfig["IdentityManagement"]["Agents"][CurrentNodeName];
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
                    Node.InnerText = Utils.Crypto.ProtectedData.EncryptString(Node.InnerText, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

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
                    returnValue = Utils.Crypto.ProtectedData.DecryptString(Node.InnerText, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);
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
                    CurrentNode[Name].InnerText = Utils.Crypto.ProtectedData.EncryptString(NewValue, System.Security.Cryptography.DataProtectionScope.CurrentUser, null);

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