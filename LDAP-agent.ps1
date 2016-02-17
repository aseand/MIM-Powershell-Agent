# FIM agent LDAP


Add-Type -Path 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\Microsoft.MetadirectoryServicesEx.dll'
. "LDAP_Client.ps1"

#Log, name, config 
$global:logger = $null
$global:MA_Name = $null
$global:Config = $null


#LDAP
$global:LdapConnection = $null
$global:SearchRequest = $null
$global:PageResultRequestControl = $null

#Schema
$global:Schema = $null
$global:SchemaObjectClassList = New-Object System.Collections.Generic.List[System.String]

#Other
$global:ImportRunStepPageSize = $null
$global:PageSize = $null
$global:DeleteObjectJob = $null
$global:DeleteList = $null

<#Initialize
#>
Function Initialize{
	Param
    (
		$logger,
		$MA_Name,
		$Config
	)
	$global:logger = $logger
	$global:MA_Name = $MA_Name
	$global:Config = $Config
	$global:logger.Info("Run Initialize")
	#TBD config
}

#region Help Functions
#Funtion MSSQLFill as string
Function MSSQLFill{
	Param
	(
		[string]$ConnectionString,
		[system.Data.DataTable]$DataTable,
		[string]$SourceTableName,
		[string]$sqlcommand
		
	)

	if(!$DataTable -OR $DataTable.HasErrors){
		write-error "DataTable empty or has errors" # | Out-File -FilePath $errorlog -Append 
		return
	}
	
	$Connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
	$Connection.Open()

	$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.Connection = $Connection
	$SqlCmd.CommandTimeout = $Connection.ConnectionTimeout
	if($sqlcommand.Length -gt 0){
		$SqlCmd.CommandText = $sqlcommand
	}
	else{
		$SqlCmd.CommandText = [string]::Format("SELECT * FROM {0}", $SourceTableName)
	}
	$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
	$SelectCount = $Adapter.Fill($DataTable)
	if ($SelectCount -eq -1){
		write-error "Can not select from " + $SourceTableName #| Out-File -FilePath $errorlog -Append 
	}
	$Adapter.Dispose()
	$SqlCmd.Dispose()
	$Connection.Close()
	$Connection.Dispose()
}

#Help Funtion GetAttibutelist for GetSchema, get all superior classes
Function GetAttibutelist{
	Param
    (
		$objectclasse,
		$objectclassesList
	)
	$List = New-Object System.Collections.ArrayList
	$object = $objectclassesList[$objectclasse]
	
	if($object.MAY.Count -gt 0){
		[void]$List.AddRange($object.MAY)
	}
	if($object.MUST.Count -gt 0){
		[void]$List.AddRange($object.MUST)
	}
	
	foreach($SUP in $object.SUP)
	{
		[void]$List.AddRange([System.Collections.ArrayList](GetAttibutelist $SUP $objectclassesList))
	}
	
	return ,$List
}
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
	$DeprovisionAction = [Microsoft.MetadirectoryServices.DeprovisionAction]::Disconnect
	
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
	
	switch(FlowRuleName)
	{
		"givenName" {$CSEntry["givenName"] = "test"}
	}
}

Function IMASynchronization.MapAttributesForExport{

	Param
    (
		$FlowRuleName,
		$MVEntry,
		$CSEntry
	)
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
	
	if ($MVEntry.ObjectType -eq "person")
	{
		if ($MVEntry["AnvandarID-Status"].Value -eq 1)
		{
			$CSentry = ManagementAgent.Connectors.StartNewConnector("inetOrgPerson", @("HSAPersonExtension","inetOrgPerson","organizationalPerson","person","securePerson","top"))
			$RDN = "cn=" + 
			$ParentDN = "ou=Ospecificerad placering,o=Landstinget Dalarna,l=Dalarnas län,c=SE"
			$CSentry.DN = 
			
		}
	}
	
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
Function IMAExtensible2GetCapabilitiesEx.GetCapabilitiesEx{

	$global:logger.debug("Start Capabilitie")
	$MACapabilities = New-Object Microsoft.MetadirectoryServices.MACapabilities
	$MACapabilities.ConcurrentOperation = $true
	$MACapabilities.DeleteAddAsReplace = $false
	$MACapabilities.DeltaImport = $true
	$MACapabilities.DistinguishedNameStyle = [Microsoft.MetadirectoryServices.MADistinguishedNameStyle]::Ldap
	$MACapabilities.ExportType = [Microsoft.MetadirectoryServices.MAExportType]::AttributeUpdate
	$MACapabilities.FullExport = $false
	$MACapabilities.NoReferenceValuesInFirstExport = $false
	$MACapabilities.ExportPasswordInFirstPass = $false
	$MACapabilities.Normalizations = [Microsoft.MetadirectoryServices.MANormalizations]::None
	$MACapabilities.ObjectRename = $true
	$MACapabilities.IsDNAsAnchor = $true
	
	$MACapabilities
	$global:logger.debug("End Capabilitie")
}

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
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("ObjectRename",$True))
				#NoReferenceValuesInFirstExport : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("NoReferenceValuesInFirstExport",$False))
				#DeltaImport                    : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("DeltaImport",$True))
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
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("ExportType",[string[]]("AttributeUpdate","AttributeReplace","ObjectReplace","MultivaluedReferenceAttributeUpdate"),$false,"ObjectReplace"))
				#Normalizations                 : None
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateDropDownParameter("Normalizations",[string[]]("None","Uppercase","RemoveAccents"),$false,"None"))
				#IsDNAsAnchor                   : False
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("IsDNAsAnchor",$False))
				#SupportImport                  : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportImport",$True))
				#SupportExport                  : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportExport",$True))
				#SupportPartitions              : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportPartitions",$True))
				#SupportPassword                : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportPassword",$True))
				#SupportHierarchy               : True
				$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SupportHierarchy",$True))
			}
		}
		"Connectivity" {
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateLabelParameter("Powershell script"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetParameters","","C:\MA-scripts\LDAP-PS-AGENT\LD-MA-Test.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetCapabilitiesEx","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetSchema","","C:\MA-scripts\LDAP-PS-AGENT\LD-MA-Test.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallImport","","C:\MA-scripts\LDAP-PS-AGENT\LD-MA-Test.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallExport","","C:\MA-scripts\LDAP-PS-AGENT\LD-MA-Test.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2Password","","C:\MA-scripts\LDAP-PS-AGENT\LD-MA-Test.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateLabelParameter("Credential (Optional)"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("User","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateEncryptedStringParameter("Password","",""))
		}
		"Global" {}
		"Partition" {}
		"RunStep" {}
		"Schema" {}
	}
	,$ConfigParameterDefinitions 
	
	$global:logger.debug("End GetConfigParameters")
}
<#
	Get LdapSchema
	return Microsoft.MetadirectoryServices.Schema
#>
Function IMAExtensible2GetSchema.GetSchema{
	Param
    (
		$ConfigParameters
	)
	try
	{
		$Schema = New-Object Microsoft.MetadirectoryServices.Schema
		#$logger.Debug("Run GetSchema")
		#
		$temp = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($configParameters["Password"].SecureValue)
		#$UnsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)
		$LdapConnection = CreateLDAPConnection -server "localhost" -port 389 -NetworkCredential (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),"")) -AuthType ([System.DirectoryServices.Protocols.AuthType]::Basic)
		
		$ldapRoot = BeginPagedSearch $LdapConnection $null "(objectClass=*)" @() 1000 ([System.DirectoryServices.Protocols.SearchScope]::Base)
		$LdapSchema = BeginPagedSearch $LdapConnection $ldapRoot.Attributes.subschemasubentry[0] "(objectClass=*)" @('ObjectClasses','attributeTypes') 1000 ([System.DirectoryServices.Protocols.SearchScope]::Base)

		$AttributesList = @{}
		$Name = $null
		$type = [Microsoft.MetadirectoryServices.AttributeType]::String
		$MulitValue = $true
		$USAGE=$null
		$AttributeOperation=[Microsoft.MetadirectoryServices.AttributeOperation]::ImportExport
		$SUP=$null
		
		#Split on schema syntax names
		$Split = $LdapSchema.Attributes.attributetypes.GetValues([System.string]) | % {$_ -csplit "\b(NAME|DESC|SYNTAX|SUP|AUX|SINGLE-VALUE|MUST|MAY|NO-USER-MODIFICATION|USAGE|ABSTRACT|STRUCTURAL|OBSOLETE|AUXILIARY|EQUALITY|SUBSTR|EWOS|ORDERING|COLLECTIVE)"}
		#populate attibute list
		for($i = 0; $i -lt $Split.Count; $i++)
		{
			switch($Split[$i])
			{
				"NAME" 					{

											if($Name -ne $null){
												$AttributesList.Add($Name,(new-object PSObject -Property @{ Type=$type; MulitValue=$MulitValue; USAGE=$USAGE; AttributeOperation=$AttributeOperation; SUP=$SUP; }))
												$Name = $null
												$type = [Microsoft.MetadirectoryServices.AttributeType]::String
												$MulitValue = $true
												$USAGE=$null
												$AttributeOperation=[Microsoft.MetadirectoryServices.AttributeOperation]::ImportExport
												$SUP=$null
											}
											$s = $Split[$i+1].Replace("'","").Replace("(","").Replace(")","").Trim().Split(" ")
											$Name = $s[0]
											$i++
										}
				#"DESC" 				{}
				"SYNTAX" 				{
											switch( ($Split[$i+1].Replace("'","").Replace("(","").Replace(")","").Trim().Split("{"))[0] )
											{	
												"1.3.6.1.4.1.1466.115.121.1.40"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.4"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.5"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.8"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.9"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.10"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.23"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.1466.115.121.1.28"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.3.6.1.4.1.4203.666.11.10.2.1" {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"1.2.840.113556.1.4.903"         {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"2.16.840.1.113719.1.1.5.1.12"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"2.16.840.1.113719.1.1.5.1.13"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
												"2.16.840.1.113719.1.1.5.1.16"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Binary}
																				
												"1.3.6.1.4.1.1466.115.121.1.7"	 {$type = [Microsoft.MetadirectoryServices.AttributeType]::Boolean}
																				
												"1.3.6.1.4.1.1466.115.121.1.1"   {$type = [Microsoft.MetadirectoryServices.AttributeType]::Reference}
												"1.3.6.1.4.1.1466.115.121.1.12"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Reference}
												"1.2.36.79672281.1.5.0"			 {$type = [Microsoft.MetadirectoryServices.AttributeType]::Reference}
																				
												"1.3.6.1.4.1.1466.115.121.1.36"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Integer}
												"1.2.840.113556.1.4.906"         {$type = [Microsoft.MetadirectoryServices.AttributeType]::Integer}
												"1.3.6.1.4.1.1466.115.121.1.27"  {$type = [Microsoft.MetadirectoryServices.AttributeType]::Integer}
												"2.16.840.1.113719.1.1.5.1.22"	 {$type = [Microsoft.MetadirectoryServices.AttributeType]::Integer}								
											}
											$i++
										}
				"SUP" 					{$SUP=$Split[$i+1].Replace(")","").Trim();$i++}
				#"AUX" 					{}
				"SINGLE-VALUE" 			{$MulitValue=$false}
				#"MUST" 				{}
				#"MAY" 					{}
				"NO-USER-MODIFICATION"	{ $AttributeOperation=[Microsoft.MetadirectoryServices.AttributeOperation]::ImportOnly;}
				"USAGE" 				{$USAGE=$Split[$i+1].Replace(")","").Trim();$i++}
				#"ABSTRACT" 			{}
				#"STRUCTURAL" 			{}
				#"OBSOLETE" 			{}
				#"EQUALITY" 			{}
				#"SUBSTR" 				{}
				#"EWOS" 				{}
				#"ORDERING" 			{}
				#"COLLECTIVE" 			{}
				default					{}
			}
		}
		#last
		if($Name -ne $null)
		{
			[void]$AttributesList.Add($Name,(new-object PSObject -Property @{ Type=$type; MulitValue=$MulitValue; USAGE=$USAGE; AttributeOperation=$AttributeOperation; SUP=$SUP; }))
		}

		#directoryOperation list 
		#$directoryOperationlist = @()
		$directoryOperationlist = New-Object System.Collections.ArrayList
		foreach($AttributesName in $AttributesList.Keys)
		{
			if( $AttributesList[$AttributesName].USAGE -eq "directoryOperation" -AND $AttributesList[$AttributesName].AttributeOperation -eq [Microsoft.MetadirectoryServices.AttributeOperation]::ImportOnly )
			{
				[void]$directoryOperationlist.Add($AttributesName)
			}
		}

		
		$objectclassesList = @{}
		$Name = $null
		$SUP = $null
		$MUST = $null
		$MAY = $null
		$ClassType = $false
		$OBSOLETE = $false
		
		#split on syta names
		$Split = $LdapSchema.Attributes.objectclasses.GetValues([System.string]) | % {$_ -csplit "\b(NAME|DESC|SYNTAX|SUP|SINGLE-VALUE|MUST|MAY|NO-USER-MODIFICATION|USAGE|ABSTRACT|STRUCTURAL|OBSOLETE|AUXILIARY|EQUALITY|SUBSTR|EWOS|ORDERING|COLLECTIVE)"}
		#populate object class list
		for($i = 0; $i -lt $Split.Count; $i++)
		{
			switch($Split[$i])
			{
				"NAME" 					{
											if(-not $OBSOLETE -AND $Name -ne $null)
											{
												$objectclassesList.Add($Name,(new-object PSObject -Property @{ ClassType=$ClassType; SUP=$SUP; MUST=$MUST; MAY=$MAY; }))
												$Name = $null
												$SUP = $null
												$MUST = $null
												$MAY = $null
												$AUXILIARY = $false
												$OBSOLETE = $false
											}
											$s = $Split[$i+1].Replace("'","").Replace("(","").Replace(")","").Trim().Split(" ")
											$Name = $s[0]
											$i++
										}
				#"DESC" 				{}
				#"SYNTAX" 				{}
				"SUP" 					{$SUP = $Split[$i+1].Trim();$i++;}
				#"AUX" 					{}
				#"SINGLE-VALUE" 		{}
				"MUST" 					{$MUST = $Split[$i+1].Replace(" ","").Replace("(","").Replace(")","").Trim().Split("`$");$i++;}
				"MAY" 					{$MAY = $Split[$i+1].Replace(" ","").Replace("(","").Replace(")","").Trim().Split("`$");$i++;}
				#"NO-USER-MODIFICATION" {}
				#"USAGE" 				{}
				"ABSTRACT" 				{$ClassType = "ABSTRACT"}
				"STRUCTURAL" 			{$ClassType = "STRUCTURAL"}
				"AUXILIARY" 			{$ClassType = "AUXILIARY"}
				"OBSOLETE" 				{$OBSOLETE = $true}
				#"EQUALITY" 			{}
				#"SUBSTR" 				{}
				#"EWOS" 				{}
				#"ORDERING" 			{}
				#"COLLECTIVE" 			{}
				default					{}
			}
		}
		#last
		if($Name -ne $null)
		{
			$objectclassesList.Add($Name,(new-object PSObject -Property @{ ClassType=$ClassType; SUP=$SUP; MUST=$MUST; MAY=$MAY; }))
		}

		#Create list of AUXILIARY classes Attibutes
		#$AUXILIARYlist = @()
		$AUXILIARYlist = New-Object System.Collections.ArrayList
		foreach($objectClassName in $objectclassesList.Keys)
		{
			if( $objectclassesList[$objectClassName].ClassType -eq "AUXILIARY" -AND $objectclassesList[$objectClassName].SUP -ne $null )
			{
				#Get all attibute, call GetAttibutelist funtion
				[void]$AUXILIARYlist.AddRange( (GetAttibutelist $objectClassName $objectclassesList) )
			}
		}
		
		#Add object to Schema
		foreach($objectClassName in $objectclassesList.Keys)
		{
			#Only STRUCTURAL object is allowed to be SchemaTypes
			if($objectclassesList[$objectClassName].ClassType -eq "STRUCTURAL")
			{
				$SchemaType = [Microsoft.MetadirectoryServices.SchemaType]::Create($objectClassName,$false)
				
				#PossibleDNComponentsForProvisioning
				foreach($DN in ($objectclassesList[$objectClassName].MUST))
				{
					$SchemaType.PossibleDNComponentsForProvisioning.Add($DN)
				}
				
				
				#MUST MAY and SUP Attributes +directoryOperationlist, call GetAttibutelist funtion
				foreach($AttributeName in (GetAttibutelist $objectClassName $objectclassesList)+$directoryOperationlist)
				{
					if(-not $SchemaType.Attributes.Contains($AttributeName) -AND $AttributeName.length -gt 0)
					{
						$AttibuteObj = $AttributesList[$AttributeName]
						
						#Mulite value
						if($AttibuteObj.MulitValue)
						{
							$AddAttribute = [Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($AttributeName, $AttibuteObj.Type, $AttibuteObj.AttributeOperation)
							$AddAttribute.HiddenByDefault = $false
							$SchemaType.Attributes.Add($AddAttribute)
						}#singel
						else
						{
							$AddAttribute = [Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($AttributeName, $AttibuteObj.Type, $AttibuteObj.AttributeOperation)
							$AddAttribute.HiddenByDefault = $false
							$SchemaType.Attributes.Add($AddAttribute)
						}
					}
				}
				
				#Add AUXILIARY Attributes, set Hidden By Default
				foreach($AttributeName in $AUXILIARYlist)
				{
					if(-not $SchemaType.Attributes.Contains($AttributeName) -AND $AttributeName.length -gt 0)
					{
						$AttibuteObj = $AttributesList[$AttributeName]

						#Mulite value
						if($AttibuteObj.MulitValue)
						{
							$AddAttribute = [Microsoft.MetadirectoryServices.SchemaAttribute]::CreateMultiValuedAttribute($AttributeName, $AttibuteObj.Type, $AttibuteObj.AttributeOperation)
							$AddAttribute.HiddenByDefault = $true
							$SchemaType.Attributes.Add($AddAttribute)
						}#singel
						else
						{
							$AddAttribute = [Microsoft.MetadirectoryServices.SchemaAttribute]::CreateSingleValuedAttribute($AttributeName, $AttibuteObj.Type, $AttibuteObj.AttributeOperation)
							$AddAttribute.HiddenByDefault = $true
							$SchemaType.Attributes.Add($AddAttribute)
						}
					}
				}
				#Add type to Schema
				$Schema.Types.Add($SchemaType)
			}
		}
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
Function IMAExtensible2CallImport.OpenImportConnection{
	Param
    (
		$ConfigParameters,
		$Schema,
		$OpenImportConnectionRunStep
	)
	try
	{
		$OpenImportConnectionResults = New-Object Microsoft.MetadirectoryServices.OpenImportConnectionResults
		$global:logger.Debug("Run IMAExtensible2CallImport.OpenImportConnection")
		$global:logger.Debug("PageSize: "+$OpenImportConnectionRunStep.PageSize)
		#InitializeConfig
	
		#$configParameters["Name"].Name
		#$configParameters["Name"].Value
		
		# OpenImportConnectionRunStep para
		#$ImportRunStepImportType = OpenImportConnectionRunStep.ImportType;
		#[Microsoft.MetadirectoryServices.OperationType):: Full Delta FullObject
		#$ImportRunStepCustomData = OpenImportConnectionRunStep.CustomData;
		#$ImportRunStepPageSize = $OpenImportConnectionRunStep.PageSize;
		$global:PageSize = $OpenImportConnectionRunStep.PageSize
		$global:Schema = $Schema
		
		$temp = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($configParameters["Password"].SecureValue)
		$global:LdapConnection = CreateLDAPConnection -server "localhost" -port 389 -NetworkCredential (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),"")) -AuthType ([System.DirectoryServices.Protocols.AuthType]::Basic)
		
		
		#$SchemaAttributsList = @()
		$SchemaAttributsList = New-Object System.Collections.Generic.List[System.String]
		#Populate attibute list and object class lists
		foreach($type in $Schema.Types)
		{
			[void]$global:SchemaObjectClassList.Add($type.Name)
			#$global:logger.Debug("Schema type: "+$type.Name)
			foreach ($attibute in $type.AnchorAttributes)
			{
				if(-not $SchemaAttributsList.Contains($attibute.Name))
				{
					[void]$SchemaAttributsList.Add($attibute.Name)
					#$global:logger.Debug("attibute: "+$attibute.Name)
				}
			}
			foreach($attibute in $type.Attributes)
			{
				if(-not $SchemaAttributsList.Contains($attibute.Name))
				{
					[void]$SchemaAttributsList.Add($attibute.Name)
					#$global:logger.Debug("attibute: "+$attibute.Name)
				}
			}
		}
		#objectclass as attibute if dont exist (top class)
		if(-not $SchemaAttributsList.Contains("objectclass"))
		{
			[void]$SchemaAttributsList.Add("objectclass")
		}
		#$global:logger.Debug("$global:SchemaObjectClassList")
		#$global:logger.Debug("$SchemaAttributsList")
		
		#Crate LDAP filter
		$filter = ""
		$LastDate = $OpenImportConnectionRunStep.CustomData
		#Full
		if($OpenImportConnectionRunStep.ImportType -eq [Microsoft.MetadirectoryServices.OperationType]::Full -OR $LastDate.length -lt 1)
		{
			$filter = "(|(objectClass="+[string]::Join(")(objectClass=",$global:SchemaObjectClassList)+"))"
		}#Delta
		elseif($OpenImportConnectionRunStep.ImportType -eq [Microsoft.MetadirectoryServices.OperationType]::Delta)
		{		
			$filter = "(&(|(objectClass="+[string]::Join(")(objectClass=",$global:SchemaObjectClassList)+"))(|(modifyTimestamp>=$LastDate)(createTimestamp>=$LastDate)))"

			$global:runspace = [runspacefactory]::CreateRunspace()
			$global:runspace.Open()
			$global:runspace.SessionStateProxy.SetVariable('MSSQLFillFunction',${function:MSSQLFill})
			$global:runspace.SessionStateProxy.SetVariable('SchemaObjectClassList',$global:SchemaObjectClassList)
			$global:runspace.SessionStateProxy.SetVariable('LdapConnection',$global:LdapConnection)
			$global:DeleteJob = [powershell]::Create()
			$global:DeleteJob.Runspace = $global:runspace
			
			#Crate backgroud jobb, delete entrys (missing CS entrys from LDAP)
			[void]$global:DeleteJob.AddScript({
				. "C:\MA-scripts\script\LDAP_Client.ps1"
				. ([ScriptBlock]::Create($MSSQLFillFunction))
				
				#Delete object
				
				#TBD get MA-name from MA dir, get GUID
				
				$SQLserver = "localhost"
				$table = New-Object system.Data.DataTable "mms_connectorspace"
				$SqlCommand =  "select object_id, rdn, pobject_id, object_type "
				$SqlCommand += "from mms_connectorspace (nolock) "
				$SqlCommand += "where ma_id = 'FF432718-31C9-4647-B23B-931CBE67B84E' "
				$SqlCommand += "and (object_type in ('" + [string]::Join("','",$SchemaObjectClassList) +"') "
				$SqlCommand += "or object_type is null )"
				$SqlCommand += "and is_obsoletion = 0 "
				$SqlCommand += "order by pobject_id"
				#Get CS entrys
				MSSQLFill -sqlcommand $SqlCommand -ConnectionString ([string]::Format("Data Source={0};Initial Catalog={1};Integrated Security=SSPI;",$SQLserver,"FIMSynchronizationService")) -DataTable $table

				$list = @{}
				$DNList = New-Object System.Collections.ArrayList
				#Create list, add parent rdn to child rdn
				foreach($row in $table)
				{
					$pobject_id = $row.pobject_id
					if($list.Contains($pobject_id))
					{
						$dn = $row.rdn+","+$list[$pobject_id]
						[void]$list.Add($row.object_id,$dn)
						if($row.object_type -ne [System.DBNull]::Value)
						{
							[void]$DNList.Add($dn)
						}
					}
					else
					{
						[void]$list.Add($row.object_id,$row.rdn)
					}
				}

				#Get Full DN list and remove DN from CS list
				$ObjectClassList = ("(|(objectClass="+[string]::Join(")(objectClass=",$SchemaObjectClassList)+"))")
				$attibutes = @("dn")
				$FullDN = BeginPagedSearch $LdapConnection $null $ObjectClassList @("dn") 1000 ([System.DirectoryServices.Protocols.SearchScope]::Subtree)
				$FullDN | % {[void]$DNList.Remove($_.DistinguishedName)}
				#return value
				$DNList
			})
			$global:DeleteJobhandle = $global:DeleteJob.BeginInvoke()
		}
		$OpenImportConnectionResults.CustomData = (get-date).ToUniversalTime().ToString("yyyyMMddhhmmssZ")
		#$global:logger.Debug($filter)
		
		#$CustomData.Clear()
		$baseDN = "c=SE"
		#$Ldaptemp = BeginPagedSearch $LdapConnection $baseDN $filter $SchemaAttributsList 1000 ([System.DirectoryServices.Protocols.SearchScope]::Subtree)
		#$global:logger.Debug("Count {0}",$Ldaptemp.Entries.count)
		$global:logger.Debug($baseDN)
		$global:logger.Debug($filter)
		##$global:logger.Debug($SchemaAttributsList)
		[void]($SchemaAttributsList | % {$global:logger.Debug($_)})
		
		#Crate SearchRequest and page cotrol
		[void]($global:SearchRequest = new-object System.DirectoryServices.Protocols.SearchRequest($baseDN, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree, $null))
		[void]($SchemaAttributsList | % {$global:SearchRequest.Attributes.Add($_)})
		[void]($global:PageResultRequestControl = new-object System.DirectoryServices.Protocols.PageResultRequestControl($OpenImportConnectionRunStep.PageSize))
		[void]($global:SearchRequest.Controls.Add($global:PageResultRequestControl))

		#$CustomData.Clear()
		#$CustomData.Append("OpenImport")
		#$global:logger.Debug("Run OpenImport end")
		$global:logger.Debug("End IMAExtensible2CallImport.OpenImportConnection")
		
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
		$CurrentCount = 0
		
		if($global:DeleteList -eq $null)
		{
			#get searchResponse 
			$searchResponse = [System.DirectoryServices.Protocols.SearchResponse] $global:LdapConnection.SendRequest($global:SearchRequest)
			#get Cookie page control
			$searchResponse.Controls | ? { $_ -is [System.DirectoryServices.Protocols.PageResultResponseControl] } | % {
				$global:PageResultRequestControl.Cookie = ([System.DirectoryServices.Protocols.PageResultResponseControl]$_).Cookie
			}
			
			#Add entry to CS entry
			#$global:logger.Debug("searchResponse {0}",$searchResponse.Entries.count)
			if($searchResponse.Entries.count -gt 0)
			{
				foreach ($Entrie in $searchResponse.Entries) 
				{ 
					$CSEntry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
					$CSEntry.ObjectModificationType = [Microsoft.MetadirectoryServices.ObjectModificationType]::Add
					$CSEntry.DN = $Entrie.DistinguishedName
					
					#Select Objectclass, first objectclass is primary, work through list first hit is the right class
					$objectclassAttribute = $Entrie.Attributes["objectclass"]
					for($i=0;$i -lt $objectclassAttribute.Count; $i++ )
					{
						if($global:SchemaObjectClassList -contains $objectclassAttribute[$i])
						{
							$CSEntry.ObjectType = $objectclassAttribute[$i]
							break
						}
					}
					
					#Add attributes
					foreach($SchemaAttibute in $Schema.Types[$CSEntry.ObjectType].Attributes)
					{
						if($SchemaAttibute.IsMultiValued)
						{
							$Valuelist =  New-Object 'System.Collections.Generic.List[Object]'
							$Attribute = $Entrie.Attributes[$SchemaAttibute.Name]
							for($i=$Attribute.count-1; $i -gt -1; $i--)
							{
								$Valuelist.add($Attribute[$i])
							}
							
							if($Valuelist.Count -gt 0)
							{
								$CSEntry.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($SchemaAttibute.Name,$Valuelist.ToArray()))
							}
						}
						else
						{
							if($Entrie.Attributes[$SchemaAttibute.Name].Count -gt 0)
							{
								$CSEntry.AttributeChanges.Add([Microsoft.MetadirectoryServices.AttributeChange]::CreateAttributeAdd($SchemaAttibute.Name,$Entrie.Attributes[$SchemaAttibute.Name][0]))
							}
						}
					}
					# add CSEntry
					$GetImportEntriesResults.CSEntries.Add($CSEntry)
				}
			}
			
			$CurrentCount = $searchResponse.Entries.count 
			$GetImportEntriesResults.MoreToImport = ($global:PageResultRequestControl.Cookie.length -gt 0)
			
			#if no more entrys try delete list by geting jobb
			if(-not $GetImportEntriesResults.MoreToImport -AND $global:DeleteJobhandle -ne $null)
			{
				While (-Not $global:DeleteJobhandle.IsCompleted) 
				{
					Start-Sleep -Milliseconds 100
				}
				$global:DeleteList = $global:DeleteJob.EndInvoke($global:DeleteJobhandle)
			}
		}
		#Work DeleteList
		if($global:DeleteList -ne $null)
		{
			while($CurrentCount -lt $global:PageSize -AND $global:DeleteList.Count -gt 0)
			{
				$CSEntry = [Microsoft.MetadirectoryServices.CSEntryChange]::Create()
				$CSEntry.ObjectModificationType = [Microsoft.MetadirectoryServices.ObjectModificationType]::Delete
				$CSEntry.DN = $global:DeleteList[0]
				$GetImportEntriesResults.CSEntries.add($CSEntry)
				
				$global:DeleteList.RemoveAt(0)
				$CurrentCount++
			}
			$GetImportEntriesResults.MoreToImport = ($global:DeleteList.Count -gt 0)
		}


		
		$GetImportEntriesResults
		
		$global:logger.Debug("{0}",$GetImportEntriesResults.CSEntries.Count)
		$global:logger.Debug("End IMAExtensible2CallImport.GetImportEntries")
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

Function IMAExtensible2CallImport.CloseImportConnection{
	Param
    (
		$CloseImportConnectionRunStep
	)
	try
	{
		$CloseImportConnectionResults = New-Object Microsoft.MetadirectoryServices.CloseImportConnectionResults
		$global:logger.Debug("Run CloseImport")
		
		if($global:DeleteJob -ne $null)
		{
			$global:runspace.Close()
			$global:DeleteJob.Dispose()
		}
		
		$CloseImportConnectionResults
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
Function IMAExtensible2CallExport.OpenExportConnection{
	Param
    (
		$ConfigParameters,
		$Schema,
		$OpenExportConnectionRunStep
	)
	try
	{
		$global:logger.Debug("Run OpenExport")
		# $exportRunStep
		# .BatchSize
		$temp = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($configParameters["Password"].SecureValue)
		$global:LdapConnection = CreateLDAPConnection -server "localhost" -port 389 -NetworkCredential (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),"")) -AuthType ([System.DirectoryServices.Protocols.AuthType]::Basic)


	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

Function IMAExtensible2CallExport.PutExportEntries{
	Param
    (
		$CSEntryChanges
	)
	try
	{
		$global:logger.Debug("Run PutExportEntries")
		$PutExportEntriesResults = New-Object Microsoft.MetadirectoryServices.PutExportEntriesResults
		foreach($CSEntry in $CSEntryChanges)
		{
			#update entry i DB or something
			# $CSEntry
			# .AnchorAttributes
			# .AttributeChanges
			# .ChangedAttributeNames
			# .DN
			# .ObjectModificationType
			# .ObjectType
			# .RDN
			
			#Set DN
			$DN = $CSEntry.DN.ToString()

			switch($CSEntry.ObjectModificationType.ToString())
			{
				#"Unconfigured" {}
				#"None" 		{}
				"Add" 			{
									#AddRequest
									$Request = new-object System.DirectoryServices.Protocols.AddRequest
									$Request.DistinguishedName = $DN
									foreach ($AttributeName in $CSEntry.ChangedAttributeNames)
									{
										$Attribute  = new-object System.DirectoryServices.Protocols.DirectoryAttribute

										foreach($Value in $CSEntry.AttributeChanges[$AttributeName].Values)
										{
											[void]$Attribute.Add($Value.ToString())
										}
										[void]$Request.Attributes.Add($Attribute)
									}
								}
				#"Replace" 		{
				#					#Move entry 
				#					$Request = new-object System.DirectoryServices.Protocols.ModifyDNRequest("oldDistinguishedName","RDN","newDN")
				#				}
				"Update" 		{
									#check if entry has new DN
									#Get from CS? most be a "old name" entry?
									
									#Try get entry in LDAP, if not exist, rename have be done
									$resualt = BeginPagedSearch $global:LdapConnection $DN $null $null 1000 ([System.DirectoryServices.Protocols.SearchScope]::Base)
									if($resualt.Count -eq 0)
									{
										#try get entry by RDN
										$RDN = $DN.Substring(0,$DN.IndexOf(","))
										$resualt = BeginPagedSearch $global:LdapConnection $null ("("+$RDN+")") $null 1000 ([System.DirectoryServices.Protocols.SearchScope]::Subtree)
										if($resualt.Count -eq 1)
										{
											#Move entry 
											$parentDN = $DN.Substring($RDN.Length+1)
											$Request = new-object System.DirectoryServices.Protocols.ModifyDNRequest($resualt[0].DistinguishedName,$parentDN,$RDN)
											
											$result = $global:LdapConnection.SendRequest($Request)
											if( $result.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success)
											{
												#error
												#write-host $result.ErrorMessage $Entry.DistinguishedName 
											}
											$Request=$null
										}
										else
										{
											#error missing entry
										}
									}

									#update entry attibutes
									$Request = new-object System.DirectoryServices.Protocols.ModifyRequest
									$Request.DistinguishedName = $DN
									foreach ($AttributeName in $CSEntry.ChangedAttributeNames)
									{
										$DirectoryModification  = new-object System.DirectoryServices.Protocols.DirectoryAttributeModification
										$DirectoryModification.Name = $AttributeName
										$DirectoryModification.Operation = [System.DirectoryServices.Protocols.DirectoryAttributeOperation]::Replace
										
										foreach($ValueChange in $CSEntry.AttributeChanges[$AttributeName].ValueChanges)
										{
											#Add
											#Delete
											#Unconfigured
											if($ValueChange.ModificationType -ne [Microsoft.MetadirectoryServices.ValueModificationType]::Delete)
											{
												[void]$DirectoryModification.Add($ValueChange.Value.ToString())
											}
										}
										[void]$Request.Modifications.Add($DirectoryModification)
										
									}
				
								}
				"Delete" 		{
									#Delete entry
									$Request = new-object System.DirectoryServices.Protocols.DeleteRequest($CSEntry.DN.ToString())
								}
			}

			#Send request
			if($Request -ne $null)
			{
				$result = $global:LdapConnection.SendRequest($Request)
				
				if( $result.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success)
				{
					write-host $result.ErrorMessage $Entry.DistinguishedName 
				}
				#add error export result to ExportEntriesResults
				#$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorInvalidAttributeValue)
				#$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
			}
			else
			{
				#error no Request for entry
			}
		}
		$PutExportEntriesResults
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
	}
}

Function IMAExtensible2CallExport.CloseExportConnection{
	Param
    (
		$CloseExportConnectionRunStep
	)
	try
	{
		$global:logger.Debug("Run CloseExport")
		#close
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
