# MIM-Powershell-Agent
Powershell agent for FIM or MIM (Microsoft Identity Manager)

Full support to run Powershell script file for all interface functions as rule extensions or full agent
Powershell script file have same functions as c# file
Exampel script is provided in example-Agents

IMASynchronization
IMVSynchronization
IMAExtensible2GetCapabilitiesEx
IMAExtensible2GetSchema
IMAExtensible2GetParameters
IMAExtensible2CallImport
IMAExtensible2CallExport
IMAExtensible2Password


# Install
Copy MIM-Powershell-Agent.dll, NLog.dll NLog.xml to MIM extensions directory
Create reg keys (MIM-Powershell-Agent.reg), update path for default loging, LoggingPath.
Optinal keys:
IMAExtensible2InitalScript - if need to use script for GetConfigParameters (new agents)
ArchiveAboveSizeMb - Nlog parameter for archive file in Mb, default 20
MaxArchiveFiles - Nlog parameter for archive files, dufalt 10
MaxLogLevel - Nlog max log lever, default Fatal
MinLogLevel - Nlog min log lever, default Info
LoggingConfigurationFullPath - To use Nlog config file insted of default, full path


# Deploy new agent
Create new agent in MIM
select extensible Connectivity 2.0 and MIM-Powershell-Agent.dll, refresh interface
Next update path to script file to use
Continue to deploy as normal agent
Log are write to LoggingPath\MIM-Powershell-Agent.log as default


# enable Metaverse Rule Extention, IMVSynchronization
Change rule extensions and browse to MIM-Powershell-Agent.dll
Create IMVSynchronization.ps1 script under any agent working directory, ex 'Synchronization Service\MaData\AD\'
Log are write to 'Synchronization Service\MaData\MIM-Powershell-Agent.log' as default


# enable Synchronization Rule, IMASynchronization
Select MIM-Powershell-Agent.dll in agent config for extensions dll.
Create IMASynchronization.ps1 script under any agent working directory, ex Synchronization Service\MaData\AD\
Log are write to 'Synchronization Service\MaData\<agent name>\MIM-Powershell-Agent.log' as default

