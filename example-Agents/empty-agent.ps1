
#Add-Type -Path 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\Microsoft.MetadirectoryServicesEx.dll'

<#
	FIM/MIM Adv.Powershell agent script 2015 Anders Åsén Landstinget Dalarna

	This script is loaded and invoke from FIM/MIM MIM Adv.Powershell agent
	Initialize function is invoke after script i loaded NOT on reloaded (on script file on change on run)
	Is recommended that global variables is to be initiated and set in initialize function, not loosely in script.
	FIM/MIM will provide Initialize with parameter: $logger, $MAName and $Config, on first load.

	If global variables is not set in Initialize function, reload may not be enable to restore those variable.
	NOTE! Don’t set ex: '$global:logger = $null' in script, as reload will load script file and then set variable that is not present after load.
	$logger will then be null until Initialize is run again.
#>


<#
	Initialize
	
	$logger is NLog instants for just this agent, NLog name will be <poweshell namespace>.<agent name> ex 'LD.IdentityManagement.Agent.Powershell.WaHSAn'
	The NLog can be use full,
	https://github.com/NLog/NLog/wiki
	
	Normal Use is : $logger.<log level>("log text") or  $logger.<log level>("{0} {1}",$variabel1,$variabel2)
	Log level: 
		Fatal
		Error
		Warn
		Info
		Debug
		Trace
		Ex: $logger.Error("This is error meassage on {0}",$UserId)
		
	$MAName is name of this current agent that is determent by FIM/MIM
	This normally use for provision or CS search
	
	$Config is extra config class in Adv.Powershell store i XML file, some extra variable is store her.
	Normal all agent have their own entry and can read and update this variables. Some variable is global (read only).
	Use get: $config["variebl name"]
	Use set: $config["variebl name"] = "value"
	Encrypt/decryp (ProtectedType) of data is done automatic. This cant only be set in use manually in XML file.
	Encrypt is using ProtectedData of current windows user in FIM/MIM
#>
Function Initialize{
	Param
    (
		$logger,
		$MAName,
		$Config
	)
	$global:logger = $logger
	$global:MAName = $MAName
	$global:Config = $Config
	$global:logger.Info("Run Initialize")
}

#region Help Functions

#endregion


#region IMASynchronization
Function IMASynchronization.Initialize{

}

Function IMASynchronization.Terminate{

}

Function IMASynchronization.ShouldProjectToMV{

	Param
    (
		$CSEntry,
		$MVObjectType
	)
	
	$result = $false
	
	$result
}

Function IMASynchronization.Deprovision{
	Param
    (
		$CSEntry
	)
	#Disconnect = 1,
    #ExplicitDisconnect = 2,
    #Delete = 3,
	#$DeprovisionAction = [Microsoft.MetadirectoryServices.DeprovisionAction]::Disconnect
	
	$DeprovisionAction
}

Function IMASynchronization.FilterForDisconnection{

	Param
    (
		$CSEntry
	)
	
	$result = $false
	
	$result
}

Function IMASynchronization.MapAttributesForJoin{

	Param
    (
		$FlowRuleName,
		$CSEntry,
		$ValueCollection
	)
	
	#switch($FlowRuleName)
	#{
	#	#"rule" { ;break}
	#}
}

Function IMASynchronization.ResolveJoinSearch{

	Param
    (
		$joinCriteriaName,
		$CSEntry,
		$rgmventry,
		$imventry,
		$MVObjectType
	)
	
	$result = $false
	
	$result
}

Function IMASynchronization.MapAttributesForImport{

	Param
    (
		$FlowRuleName,
		$CSEntry,
		$MVEntry
	)
	
	#switch($FlowRuleName)
	#{
	#	#"rule" { ;break}
	#}
}

Function IMASynchronization.MapAttributesForExport{

	Param
    (
		$FlowRuleName,
		$MVEntry,
		$CSEntry
	)
	
	#switch($FlowRuleName)
	#{
	#	#"rule" { ;break}
	#}
}
#endregion

#region IMVSynchronization
Function IMVSynchronization.Initialize{

}

Function IMVSynchronization.Terminate{

}

Function IMVSynchronization.Provision{

	Param
    (
		$MVEntry
	)
}

Function IMVSynchronization.ShouldDeleteFromMV{

	Param
    (
		$CSEntry,
		$MVEntry
	)
	
	$result = $false
	
	return $result
}
#endregion


#region IMAExtensible2
<#
	IMAExtensible2
	Only use for FIM 2010 capability
	FIM 2010 R2 and above use IMAExtensible2GetParameters.GetConfigParameters
	You may use this if config powershell agent to use it.
#>
Function IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx{

	$global:logger.debug("Start Capabilitie")
	$MACapabilities = New-Object Microsoft.MetadirectoryServices.MACapabilities
	$MACapabilities.ConcurrentOperation = $true
	$MACapabilities.DeleteAddAsReplace = $false
	$MACapabilities.DeltaImport = $false
	$MACapabilities.DistinguishedNameStyle = [Microsoft.MetadirectoryServices.MADistinguishedNameStyle]::Generic
	$MACapabilities.ExportType = [Microsoft.MetadirectoryServices.MAExportType]::AttributeUpdate
	$MACapabilities.FullExport = $false
	$MACapabilities.NoReferenceValuesInFirstExport = $false
	$MACapabilities.ExportPasswordInFirstPass = $false
	$MACapabilities.Normalizations = [Microsoft.MetadirectoryServices.MANormalizations]::None
	$MACapabilities.ObjectRename = $true
	$MACapabilities.IsDNAsAnchor = $true
	$MACapabilities.SupportImport = $true
	$MACapabilities.SupportExport = $true
	
	$MACapabilities
	$global:logger.debug("End Capabilitie")
}

<#
	IMAExtensible2GetParameters.GetConfigParameters
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2getparameters.getconfigparameters%28v=vs.100%29.aspx
	
	Config setup for agent (only use ones on create)
	
	Capabilities:
	Pre. def. variables, this may be change in FIM/MIM
	Add extra config(FIM/MIM ConfigParameters variable) if need
	
	Connectivity:
	Powershell script path normally this script file shod be use 
	Add extra config(FIM/MIM ConfigParameters variable) if need
	
	Global:
	FIM/MIM global variables
	
	Partition:
	FIM/MIM Partition variables, use if LDAP structure is use
	
	RunStep:
	FIM/MIM set def. RunStep
	
	RunStep:
	FIM/MIM set def. Schema
#>
Function IMAExtensible2GetParameters.GetConfigParameters{
	Param
    (
		$ConfigParameters,
		$ConfigParameterPage
	)
	$global:logger.debug("Start GetConfigParameters")
	$ConfigParameterDefinitions = New-Object System.Collections.Generic.List[Microsoft.MetadirectoryServices.ConfigParameterDefinition]
	$global:logger.debug($ConfigParameterPage)
	
	switch($ConfigParameterPage)
	{
		"Capabilities" {
			$global:logger.debug($configParameters["IMAExtensible2GetCapabilitiesEx"].Value.Length)
			if ($configParameters["IMAExtensible2GetCapabilitiesEx"].Value.Length -eq 0)
			{
				#DistinguishedNameStyle         : Generic
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("DistinguishedNameStyle",[string[]]("Generic","Ldap","None"),$false,"Generic"))
				#ObjectRename                   : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("ObjectRename",$False))
				#NoReferenceValuesInFirstExport : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("NoReferenceValuesInFirstExport",$False))
				#DeltaImport                    : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("DeltaImport",$False))
				#ConcurrentOperation            : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("ConcurrentOperation",$True))
				#DeleteAddAsReplace             : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("DeleteAddAsReplace",$True))
				#ExportPasswordInFirstPass      : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("ExportPasswordInFirstPass",$False))
				#FullExport                     : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("FullExport",$False))
				#ObjectConfirmation             : Normal
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("ObjectConfirmation",[string[]]("Normal","NoDeleteConfirmation","NoAddAndDeleteConfirmation"),$false,"Normal"))
				#ExportType                     : ObjectReplace
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("ExportType",[string[]]("AttributeUpdate","AttributeReplace","ObjectReplace","MultivaluedReferenceAttributeUpdate"),$false,"AttributeUpdate"))
				#Normalizations                 : None
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("Normalizations",[string[]]("None","Uppercase","RemoveAccents"),$false,"None"))
				#IsDNAsAnchor                   : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("IsDNAsAnchor",$True))
				#SupportImport                  : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportImport",$True))
				#SupportExport                  : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportExport",$True))
				#SupportPartitions              : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportPartitions",$False))
				#SupportPassword                : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportPassword",$False))
				#SupportHierarchy               : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportHierarchy",$False))
			}
			break
		}
		"Connectivity" {
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateLabelParameter("Powershell script"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetParameters","","E:\MIM-PS-agents\SharePointSite\LD-MA-SharePointSite.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetCapabilitiesEx","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetSchema","","E:\MIM-PS-agents\SharePointSite\LD-MA-SharePointSite.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallImport","","E:\MIM-PS-agents\SharePointSite\LD-MA-SharePointSite.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallExport","","E:\MIM-PS-agents\SharePointSite\LD-MA-SharePointSite.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2Password","","E:\MIM-PS-agents\SharePointSite\LD-MA-SharePointSite.ps1"))
			break
		}
		"Global" {break}
		"Partition" {break}
		"RunStep" {break}
		"Schema" {break}
	}
	,$ConfigParameterDefinitions 
	
	$global:logger.debug("End GetConfigParameters")
}

<#
	GetSchema
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2getschema.getschema(v=vs.100).aspx
	
	Setup of schema.
	Rek. to use dynamic way to generate scheme from DB/webservice/file ...
	Can crate different object type.
	If use Reference value use be sure that all ref. object is imported.
#>
Function IMAExtensible2GetSchema.GetSchema{
	Param
    (
		$ConfigParameters
	)
	try
	{
		$Schema = New-Object Microsoft.MetadirectoryServices.Schema
		$logger.Debug("Run GetSchema")

		$SchemaType = [Microsoft.MetadirectoryServices.SchemaType]::Create("Object type",$false)
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute string", [Microsoft.MetadirectoryServices.AttributeType]::String)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Integer", [Microsoft.MetadirectoryServices.AttributeType]::Integer)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Reference", [Microsoft.MetadirectoryServices.AttributeType]::Reference)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Binary", [Microsoft.MetadirectoryServices.AttributeType]::Binary)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Boolean", [Microsoft.MetadirectoryServices.AttributeType]::Boolean)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute("Multi value attribute string", [Microsoft.MetadirectoryServices.AttributeType]::String)))
		$Schema.Types.Add($SchemaType)
		
		$SchemaType = [Microsoft.MetadirectoryServices.SchemaType]::Create("some more Object type",$false)
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute string", [Microsoft.MetadirectoryServices.AttributeType]::String)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Integer", [Microsoft.MetadirectoryServices.AttributeType]::Integer)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Reference", [Microsoft.MetadirectoryServices.AttributeType]::Reference)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Binary", [Microsoft.MetadirectoryServices.AttributeType]::Binary)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute("Attribute Boolean", [Microsoft.MetadirectoryServices.AttributeType]::Boolean)))
		$SchemaType.Attributes.Add(([Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute("Multi value attribute string", [Microsoft.MetadirectoryServices.AttributeType]::String)))
		$Schema.Types.Add($SchemaType)
		
		#Return
		$Schema
	}
	Catch
	{
		$_
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

<#
	ValidateConfigParameters
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2getparameters.validateconfigparameters%28v=vs.100%29.aspx
	
	If some extra variable validation is need
#>
Function IMAExtensible2GetParameters.ValidateConfigParameters{
	Param
    (
		$ConfigParameters,
		$ConfigParameterPage
	)
	$global:logger.Debug("ValidateConfigParameters")
	$ParameterValidationResult = New-Object Microsoft.MetadirectoryServices.ParameterValidationResult
	
	$ParameterValidationResult
}
#endregion

#region Import
<#
	OpenImportConnection
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callimport.openimportconnection.aspx
	
	Open data stream connection for import,
	May read data but are not required at this point.
	
	OperationType most be followed (Full Delta)
	
	Any global variable that is need in GetImportEntries may be set her.
	NOTE! PageSize and schema are very imported to use as global variables.
#>
Function IMAExtensible2CallImport.OpenImportConnection{
	Param
    (
		$ConfigParameters,
		$Schema,
		$OpenImportConnectionRunStep
	)
	Try{
		$OpenImportConnectionResults = New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults
		$global:logger.Debug("Run IMAExtensible2CallImport.OpenImportConnection")
		$global:logger.Debug("PageSize: {0}",$OpenImportConnectionRunStep.PageSize)
		#InitializeConfig

		#$configParameters["Name"].Name
		#$configParameters["Name"].Value

		# OpenImportConnectionRunStep para
		#$ImportRunStepImportType = OpenImportConnectionRunStep.ImportType;
		#[Microsoft.MetadirectoryServices.OperationType]:: Full Delta FullObject
		#$ImportRunStepCustomData = OpenImportConnectionRunStep.CustomData;
		#$ImportRunStepPageSize = $OpenImportConnectionRunStep.PageSize;
		$global:PageSize = $OpenImportConnectionRunStep.PageSize
		$global:Schema = $Schema


			
		$global:logger.Debug("Run OpenImport end")
		$global:logger.Debug("End IMAExtensible2CallImport.OpenImportConnection")
		#$OpenImportConnectionResults.CustomData = "data"

		$OpenImportConnectionResults
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

<#
	GetImportEntries
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callimport.getimportentries.aspx
	
	This function is call multiple time, until no more data left to import
	
	You most respect PageSize and schema (attributes) for import 
	MoreToImport indicate if more data is available.
	
	Is rek. that any error should halt import stage (throw error)
	Is not rek. to use multitaskning/thrading as it may cause unforeseen problem.
#>
Function IMAExtensible2CallImport.GetImportEntries{
	Param
    (
		$GetImportEntriesRunStep
	)
	try
	{
		$global:logger.Debug("Start IMAExtensible2CallImport.GetImportEntries")

		$GetImportEntriesResults = New-Object Microsoft.MetadirectoryServices.GetImportEntriesResults
		$GetImportEntriesResults.CSEntries = New-Object System.Collections.Generic.List[Microsoft.MetadirectoryServices.CSEntryChange]
		$logger.debug("PageSize {0}",$global:PageSize)
		
		while($GetImportEntriesResults.CSEntries.Count -lt $global:PageSize -AND "more DataExist?"){
		
			#Schema check to include ObjectType
			if($Global:Schema.Types.Contains("dataEntry ObjectType")) { 
		
				$CSEntry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
				$CSEntry.ObjectModificationType = [Microsoft.MetadirectoryServices.ObjectModificationType]::Add
				$CSEntry.ObjectType = "dataEntry ObjectType"
				$CSEntry.DN = "dataEntry DN"
				
				#Add singel value Attribute
				#Schema check to include Attribute
				if($Global:Schema.Types["dataEntry ObjectType"].Attributes.Contains("Attribute name")) { 
					$CSEntry.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd("Attribute name","Attribute value"))
				}
				
				#Add Multivalue Attribute
				if($Global:Schema.Types["dataEntry ObjectType"].Attributes.Contains("Multi Attribute name")) { 
					$Valuelist =  New-Object 'System.Collections.Generic.List[Object]'
					$Valuelist.Add("Some value")
					$CSEntry.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd("Multi Attribute name",$Valuelist))
				}
				
				#Alt.
				foreach($attibute in $Global:Schema.Types["dataEntry ObjectType"].Attributes){
				
					$Value = getValueof($attibute.Name)
					
					if($Value){
						$CSEntry.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($attibute.Name,$Value))
					}
				}
				
				
				$GetImportEntriesResults.CSEntries.Add($CSEntry)
			}
		}


		$GetImportEntriesResults.MoreToImport = #$true?$false
		
		#
		$GetImportEntriesResults
		
		$global:logger.Debug("CSEntries count: {0}",$GetImportEntriesResults.CSEntries.Count)
		$global:logger.Debug("End IMAExtensible2CallImport.GetImportEntries")
	}
	Catch
	{
		$_
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

<#
	CloseImportConnection
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callimport.closeimportconnection.aspx
	
	Clean up process for closing data stream ect.
#>
Function IMAExtensible2CallImport.CloseImportConnection{
	Param
    (
		$CloseImportConnectionRunStep
	)
	try
	{
		$global:logger.Debug("Run IMAExtensible2CallImport.CloseImportConnection")
		$CloseImportConnectionResults = New-Object Microsoft.MetadirectoryServices.CloseImportConnectionResults
		#clean up

		$CloseImportConnectionResults
		$global:logger.Debug("End IMAExtensible2CallImport.CloseImportConnection")
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}
#endregion

#region export
<#
	OpenExportConnection
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callexport.openexportconnection.aspx
	
	Open data stream for export
	BatchSize is importen to use
#>
Function IMAExtensible2CallExport.OpenExportConnection{
	Param
    (
		$ConfigParameters,
		$Schema,
		$OpenExportConnectionRunStep
	)
	try
	{
		$global:logger.Debug("Run IMAExtensible2CallExport.OpenExportConnection")
		# $exportRunStep
		$Global:BatchSize = $exportRunStep.BatchSize
		$global:logger.Debug("End IMAExtensible2CallExport.OpenExportConnection")
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

<#
	PutExportEntries
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callexport.putexportentries.aspx
	
	Export data from FIM/MIM
	This function is call ones
	
	use PutExportEntriesResults to indicate error
	May stop export by throw error
#>
Function IMAExtensible2CallExport.PutExportEntries{
	Param
    (
		$CSEntryChanges
	)
	$global:logger.Debug("Run IMAExtensible2CallExport.PutExportEntries")
	
	$PutExportEntriesResults = New-Object Microsoft.MetadirectoryServices.PutExportEntriesResults
	foreach($CSEntry in $CSEntryChanges)
	{
		#Do add/delete/change
		#$CSEntry.Identifier - MIM ID
		#$CSEntry.DN - entrys Distinguished Name
		#$CSEntry.RDN - entrys Relative Distinguished Names
		#$CSEntry.ObjectType - entrys object type
		#$CSEntry.ObjectModificationType - entrys [Microsoft.MetadirectoryServices.ObjectModificationType]::Unconfigured/None/Add/Replace/Update/Delete
		#$CSEntry.AnchorAttributes - entrys Anchors as KeyedCollection  
			# Microsoft.MetadirectoryServices.AnchorAttribute
			# AttributeType DataType
			# string Name
			# object Value
		#$CSEntry.AttributeChanges - entrys Change Attribute as KeyedCollection 
			# Microsoft.MetadirectoryServices.AttributeChange
			# AttributeType DataType
			# bool IsMultiValued
			# AttributeModificationType ModificationType
			# string Name
			# IList<ValueChange> ValueChanges
		#$CSEntry.ChangedAttributeNames - entrys Change Attribute Name as IList<string>
		
		foreach($name in $CSEntry.ChangedAttributeNames){
			#update ?
			#[Microsoft.MetadirectoryServices.ValueModificationType]::Unconfigured/Add/Delete
			#$CSEntry.AttributeChanges[$name].ValueChanges[0]ModificationType
			$value = $CSEntry.AttributeChanges[$name].ValueChanges[0].Value
		}
		
		#comit? $Global:BatchSize
		#on Success or Error
		#[Microsoft.MetadirectoryServices.MAExportError]::Success ... ExportErrorConnectedDirectoryError
		#$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorConnectedDirectoryError)
		#$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
	
	}
	$PutExportEntriesResults
	$global:logger.Debug("End IMAExtensible2CallExport.PutExportEntries")
}

<#
	CloseExportConnection
	https://msdn.microsoft.com/en-us/library/microsoft.metadirectoryservices.imaextensible2callexport.closeexportconnection.aspx

	Clean up process for closing data stream ect.
#>
Function IMAExtensible2CallExport.CloseExportConnection{
	Param
    (
		$CloseExportConnectionRunStep
	)
	try
	{
		$global:logger.Debug("Run IMAExtensible2CallExport.CloseExportConnection")
		#Clean up
		$global:logger.Debug("End IMAExtensible2CallExport.CloseExportConnection")
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}
#endregion

#region password

Function IMAExtensible2Password.OpenPasswordConnection{
	Param
    (
		$ConfigParameters,
		$Partition
	)
	$ParameterValidationResult = New-Object Microsoft.MetadirectoryServices.ParameterValidationResult
	
	$ParameterValidationResult
}

Function IMAExtensible2Password.GetConnectionSecurityLevel{
	$ConnectionSecurityLevel = New-Object Microsoft.MetadirectoryServices.ConnectionSecurityLevel
	
	$ConnectionSecurityLevel
}

Function IMAExtensible2Password.ChangePassword{
	Param
    (
		$CSEntry,
		$oldPassword,
		$newPassword
	)

}

Function IMAExtensible2Password.SetPassword{
	Param
    (
		$CSEntry,
		$newPassword,
		$PasswordOptions
	)

}

Function IMAExtensible2Password.ClosePasswordConnection{

}
#endregion