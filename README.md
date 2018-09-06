# MIM-Powershell-Agent
Powershell agent for FIM or MIM (Microsoft Identity Manager)

Full support to run Powershell script file for all interface functions
Powershell script file have same functions as c# file

IMASynchronization
IMVSynchronization
IMAExtensible2GetCapabilitiesEx
IMAExtensible2GetSchema
IMAExtensible2GetParameters
IMAExtensible2CallImport
IMAExtensible2CallExport
IMAExtensible2Password


Require config file to initialize IMAExtensible2GetParameters on create new agent
Optional, Nlog can be remove from C# and Powershell script
Optional, Metaverse Rule Extention and Synchronization Rule can be use with other agent.


# Install
Copy MIM-Powershell-Agent.dll, NLog.dll NLog.xml to MIM extensions directory
ex C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\Extensions
Create conf direcory under MIM Synchronization (C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\conf)
Copy or crate xml config, IdentityManagement.Agents.Config.xml
Create nlog config, update IdentityManagement.Agents.Config.xml whit full path.


# Deploy new agent (Optional)
Create new agent tag(agent name) in IdentityManagement.Agents.Config.xml file.
Change the IMAExtensible2GetParameters in powershell agent to full path for powershell script file (new agent file).
Create new agent in MIM (most have same name as in config file), select extensible Connectivity 2.0 and MIM-Powershell-Agent.dll, refresh interface.
Next update path to script file if needed.
Continue to deploy as normal agent.


# enable Metaverse Rule Extention, IMVSynchronization (Optional)
Change enable and browse to MIM-Powershell-Agent.dll
Update IdentityManagement.Agents.Config.xml IMVSynchronization-MA-List value under Powershell agent.
All agent in the IMVSynchronization-MA-List most have value for IMVSynchronization set to full path for powershell script file.


# enable Synchronization Rule, IMASynchronization (Optional)
Select MIM-Powershell-Agent.dll in agent config, agent most have value for IMASynchronization set to full path for powershell script file.

