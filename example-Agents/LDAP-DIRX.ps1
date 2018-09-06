# FIM agent DIRX LDAP


Add-Type -Path 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\Microsoft.MetadirectoryServicesEx.dll'
. "c:\script\LDAP-Client.ps1"

#Log, name, config 
$global:logger = $null
$global:MAName = $null
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
$global:DeleteJobhandle = $null
$global:DeleteList = $null

<#Initialize
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
	#$global:logger.Info("Run Initialize")
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
			break
		}
		"Connectivity" {
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateLabelParameter("Powershell script"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetParameters","","c:\script\LDAP-DIRX.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetCapabilitiesEx","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2GetSchema","","c:\script\LDAP-DIRX.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallImport","","c:\script\LDAP-DIRX.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2CallExport","","c:\script\LDAP-DIRX.ps1"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("IMAExtensible2Password","","c:\script\LDAP-DIRX.ps1"))
			#$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateLabelParameter("Credential (Optional)"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("User","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateEncryptedStringParameter("Password","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("Server","",""))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateStringParameter("Port","","389"))
			$ConfigParameterDefinitions.Add([Microsoft.MetadirectoryServices.ConfigParameterDefinition]::CreateCheckBoxParameter("SSL",$False))
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
		$logger.Debug("Run GetSchema")
		
		# create config from $configParameters?
		#
		$temp = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($configParameters["Password"].SecureValue)
		$Params = @{            
			Server = $configParameters["Server"].Value       
			Port = $configParameters["Port"].Value       
			NetworkCredential = (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),""))
			AuthType = ([System.DirectoryServices.Protocols.AuthType]::Basic)			
		}
		if( $configParameters["SSL"].Value -eq "1"){ 
		#	$logger.Debug("SSL {0}",$configParameters["SSL"].Value)
			$Params.add("SSL",$true) 
		}
		
		$LdapConnection = CreateLDAPConnection @Params
		$global:logger.Debug($LdapConnection.Directory.Server)
		$LdapConnection.Bind()
		$global:logger.Debug($LdapConnection.SessionOptions.HostName)

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
											break
										}
				#"DESC" 				{break}
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
											break
										}
				"SUP" 					{$SUP=$Split[$i+1].Replace(")","").Trim();$i++;break}
				#"AUX" 					{break}
				"SINGLE-VALUE" 			{$MulitValue=$false;break}
				#"MUST" 				{break}
				#"MAY" 					{break}
				"NO-USER-MODIFICATION"	{ $AttributeOperation=[Microsoft.MetadirectoryServices.AttributeOperation]::ImportOnly;break}
				"USAGE" 				{$USAGE=$Split[$i+1].Replace(")","").Trim();$i++;break}
				#"ABSTRACT" 			{break}
				#"STRUCTURAL" 			{break}
				#"OBSOLETE" 			{break}
				#"EQUALITY" 			{break}
				#"SUBSTR" 				{break}
				#"EWOS" 				{break}
				#"ORDERING" 			{break}
				#"COLLECTIVE" 			{break}
				default					{break}
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
											break
										}
				#"DESC" 				{break}
				#"SYNTAX" 				{break}
				"SUP" 					{$SUP = $Split[$i+1].Trim();$i++;break}
				#"AUX" 					{break}
				#"SINGLE-VALUE" 		{break}
				"MUST" 					{$MUST = $Split[$i+1].Replace(" ","").Replace("(","").Replace(")","").Trim().Split("`$");$i++;break}
				"MAY" 					{$MAY = $Split[$i+1].Replace(" ","").Replace("(","").Replace(")","").Trim().Split("`$");$i++;break}
				#"NO-USER-MODIFICATION" {break}
				#"USAGE" 				{break}
				"ABSTRACT" 				{$ClassType = "ABSTRACT";break}
				"STRUCTURAL" 			{$ClassType = "STRUCTURAL";break}
				"AUXILIARY" 			{$ClassType = "AUXILIARY";break}
				"OBSOLETE" 				{$OBSOLETE = $true;break}
				#"EQUALITY" 			{break}
				#"SUBSTR" 				{break}
				#"EWOS" 				{break}
				#"ORDERING" 			{break}
				#"COLLECTIVE" 			{break}
				default					{break}
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
		#$global:logger.Debug("Run IMAExtensible2CallImport.OpenImportConnection")
		#$global:logger.Debug("PageSize: "+$OpenImportConnectionRunStep.PageSize)
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
		$Params = @{            
			Server = $configParameters["Server"].Value       
			Port = $configParameters["Port"].Value       
			NetworkCredential = (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),""))
			AuthType = ([System.DirectoryServices.Protocols.AuthType]::Basic)			
		}
		if( $configParameters["SSL"].Value -eq "1"){ 
		#	$logger.Debug("SSL {0}",$configParameters["SSL"].Value)
			$Params.add("SSL",$true) 
		}
		
		$global:LdapConnection = CreateLDAPConnection @Params
		#$global:LdapConnection = CreateLDAPConnection -server "localhost" -port 389 -NetworkCredential (new-object System.Net.NetworkCredential($global:Config["hsauser"],$global:Config["hsauserpwd"],"")) -AuthType ([System.DirectoryServices.Protocols.AuthType]::Basic)

		
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
		$NewDate = (get-date).ToUniversalTime().ToString("yyyyMMddhhmmssZ")
		
		#Full
		if($OpenImportConnectionRunStep.ImportType -eq [Microsoft.MetadirectoryServices.OperationType]::Full -OR $LastDate.length -lt 1)
		{
			$filter = "(|(objectClass="+[string]::Join(")(objectClass=",$global:SchemaObjectClassList)+"))"
		}#Delta
		elseif($OpenImportConnectionRunStep.ImportType -eq [Microsoft.MetadirectoryServices.OperationType]::Delta)
		{		
			$filter = "(&(|(objectClass="+[string]::Join(")(objectClass=",$global:SchemaObjectClassList)+"))(|(modifyTimestamp>=$LastDate)(createTimestamp>=$LastDate)))"

			$global:DeleteJob = [powershell]::Create()
			$global:DeleteJob.runspace = [runspacefactory]::CreateRunspace()
			$global:DeleteJob.runspace.Open()
			#$global:DeleteJob.runspace.SessionStateProxy.SetVariable('MSSQLFillFunction',${function:MSSQLFill})
			$global:DeleteJob.runspace.SessionStateProxy.SetVariable('SchemaObjectClassList',$global:SchemaObjectClassList)
			$global:DeleteJob.runspace.SessionStateProxy.SetVariable('LdapConnection',(CreateLDAPConnection @Params))
			$global:DeleteJob.runspace.SessionStateProxy.SetVariable('MAName',$global:MaName)
			$global:DeleteJob.runspace.SessionStateProxy.SetVariable('logger',$global:logger)
			$global:DeleteJob.runspace.SessionStateProxy.SetVariable('ConnectionString',$global:config["FIM-ConnectionString"])
			
			#Crate backgroud jobb, delete entrys (missing CS entrys from LDAP)
			[void]$global:DeleteJob.AddScript({
				#return
				try{
					. "c:\script\LDAP-Client.ps1"
					#. ([ScriptBlock]::Create($MSSQLFillFunction))
					
					#Delete object
					$StartTime = get-date
					
					$SqlCommand =  "select object_id, rdn, pobject_id, object_type "
					$SqlCommand += "from mms_connectorspace (nolock) "
					$SqlCommand += ("where ma_id = (select ma_id from mms_management_agent (nolock) where ma_name = '{0}' ) " -f $MAName)
					$SqlCommand += "and (object_type in ('" + [string]::Join("','",$SchemaObjectClassList) +"') "
					$SqlCommand += "or object_type is null )"
					$SqlCommand += "and is_obsoletion = 0 "
					$SqlCommand += "order by pobject_id"
					#$logger.Debug("SqlCommand:{0}",$SqlCommand)
					
					$Connection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
					$Connection.Open()
					$SqlCmd = New-Object System.Data.SqlClient.SqlCommand($SqlCommand,$Connection)
					$Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCmd
					$table = New-Object system.Data.DataTable "mms_connectorspace"
					$SelectCount = $Adapter.Fill($table)
					$Adapter.Dispose()
					$SqlCmd.Dispose()
					$Connection.Close()
					$Connection.Dispose()
					#$logger.Debug("SQL count:{0}",$table.Rows.Count)

					$temp = ""
					$list = New-Object 'system.collections.generic.dictionary[string,string]'
					$DNList = New-Object System.Collections.Generic.HashSet[string]
					#Create list, add parent rdn to child rdn
					foreach($row in $table)
					{
						if($list.TryGetValue($row.pobject_id,[ref]$temp))
						{
							$dn = $row.rdn+","+$temp
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
					$table.Dispose()

					#Get Full DN list and remove DN from CS list
					$ObjectClassList = ("(|(objectClass="+[string]::Join(")(objectClass=",$SchemaObjectClassList)+"))")
					$attibutes = @("dn")
					$LDAPlist = BeginPagedSearch $LdapConnection $null $ObjectClassList @("dn") 1000 ([System.DirectoryServices.Protocols.SearchScope]::Subtree)
					#$logger.debug("LDAP DN count:{0}",$LDAPlist.Count)
					$LDAPlist | % {[void]$DNList.Remove($_.DistinguishedName)}
					#return value
					$DNList
					#$logger.debug("Delete run time:{0}ms list count:{1}",((get-date)-$StartTime).TotalMilliseconds,$DNList.Count)
				}
				Catch
				{
					$global:logger.Error($_.Exception.Message)
					$global:logger.Error($_.Exception.Source)
					$global:logger.Error($_.Exception.StackTrace)
					$global:logger.Error($_.InvocationInfo.ScriptLineNumber)
				}
			})
			$global:DeleteJobhandle = $global:DeleteJob.BeginInvoke()
		}
		$OpenImportConnectionResults.CustomData = $NewDate
		$config["LastDate"] = $NewDate
		#$global:logger.Debug($filter)
		
		#$CustomData.Clear()
		$baseDN = "c=SE"
		#$Ldaptemp = BeginPagedSearch $LdapConnection $baseDN $filter $SchemaAttributsList 1000 ([System.DirectoryServices.Protocols.SearchScope]::Subtree)
		#$global:logger.Debug("Count {0}",$Ldaptemp.Entries.count)
		#$global:logger.Debug("BaseDN: {0}",$baseDN)
		#$global:logger.Debug("Filter: {0}",$filter)
		#$global:logger.Debug("AttributsList: {0}", ($SchemaAttributsList -join ","))
		
		#Crate SearchRequest and page cotrol
		[void]($global:SearchRequest = new-object System.DirectoryServices.Protocols.SearchRequest($baseDN, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree, $null))
		[void]($SchemaAttributsList | % {$global:SearchRequest.Attributes.Add($_)})
		[void]($global:PageResultRequestControl = new-object System.DirectoryServices.Protocols.PageResultRequestControl($OpenImportConnectionRunStep.PageSize))
		[void]($global:SearchRequest.Controls.Add($global:PageResultRequestControl))

		#$CustomData.Clear()
		#$CustomData.Append("OpenImport")
		#$global:logger.Debug("Run OpenImport end")
		#$global:logger.Debug("End IMAExtensible2CallImport.OpenImportConnection")
		
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
		#$global:logger.Debug("Start IMAExtensible2CallImport.GetImportEntries")

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
					#$logger.Debug("Wating on DeleteJobhandle...")
					Start-Sleep -Milliseconds 100
				}
				$global:DeleteList = $global:DeleteJob.EndInvoke($global:DeleteJobhandle)
				#$logger.Debug("List {0}",$global:DeleteList.Count)
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
		
		#$global:logger.Debug("{0}",$GetImportEntriesResults.CSEntries.Count)
		#$global:logger.Debug("End IMAExtensible2CallImport.GetImportEntries")
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
		#$global:logger.Debug("Run CloseImport")
		
		if($global:DeleteJob -ne $null)
		{
			$global:DeleteJob.runspace.Close()
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
		$Params = @{            
			Server = $configParameters["Server"].Value       
			Port = $configParameters["Port"].Value       
			NetworkCredential = (new-object System.Net.NetworkCredential($configParameters["User"].Value, ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($temp)),""))
			AuthType = ([System.DirectoryServices.Protocols.AuthType]::Basic)			
		}
		if( $configParameters["SSL"].Value -eq "1"){ 
		#	$logger.Debug("SSL {0}",$configParameters["SSL"].Value)
			$Params.add("SSL",$true) 
		}
		
		$global:LdapConnection = CreateLDAPConnection @Params
		#$logger.Debug($global:LdapConnection.Directory.Servers)

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
	#$logger.Debug("Run PutExportEntries")
	$PutExportEntriesResults = New-Object Microsoft.MetadirectoryServices.PutExportEntriesResults
	try
	{
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
			
			$DN_ModifyRequest = $false
			#Set DN
			$DN = $CSEntry.DN.ToString()
			#$logger.Debug("{0} | {1}",$DN,$CSEntry.ObjectModificationType.ToString())
			switch($CSEntry.ObjectModificationType.ToString())
			{
				#"Unconfigured" {break}
				#"None" 		{break}
				"Add" 			{
									#AddRequest
									$Request = new-object System.DirectoryServices.Protocols.AddRequest
									$Request.DistinguishedName = $DN
									#$logger.Debug("ChangedAttributeNames:{0}",[string]::Join(",",$CSEntry.ChangedAttributeNames))
									foreach ($AttributeName in $CSEntry.ChangedAttributeNames)
									{
										#$Attribute  = new-object System.DirectoryServices.Protocols.DirectoryAttribute
										$DirectoryModification  = new-object System.DirectoryServices.Protocols.DirectoryAttributeModification
										$DirectoryModification.Name = $AttributeName
										$DirectoryModification.Operation = [System.DirectoryServices.Protocols.DirectoryAttributeOperation]::Add

										foreach($ValueChange in $CSEntry.AttributeChanges[$AttributeName].ValueChanges){
											#$logger.Debug("{0}:{1}",$AttributeName,$ValueChange.Value.ToString())
											if($AttributeName -ne "DN"){
												[void]$DirectoryModification.Add($ValueChange.Value.ToString())
											}
										}
										[void]$Request.Attributes.Add($DirectoryModification)
									}
									break
								}
				"Replace" 		{
									#Replace entry attibutes
									$Request = new-object System.DirectoryServices.Protocols.ModifyRequest
									$Request.DistinguishedName = $DN
									#$logger.Debug("Replace on :{0}",$DN)
									#$logger.Debug("ChangedAttributeNames:{0}",[string]::Join(",",$CSEntry.ChangedAttributeNames))
									foreach ($AttributeName in $CSEntry.ChangedAttributeNames)
									{
										#Move Entry
										if($AttributeName -eq "DN")
										{
											$NewDN = $CSEntry.AttributeChanges[$AttributeName].ValueChanges[0].Value.ToString()
											$Index = $NewDN.IndexOf(",")
											

											$DNRequest = new-object System.DirectoryServices.Protocols.ModifyDNRequest
											$DNRequest.DistinguishedName = $DN
											$DNRequest.NewName  = $NewDN.Substring(0,$Index)
											$DNRequest.NewParentDistinguishedName = $NewDN.Substring($Index+1)
											#$logger.Debug("ModifyDNRequest `nDN: {0} `nNewDN: {1} `nNewParentDistinguishedName: {2} `nNewName: {3}",$DNRequest.DistinguishedName,$NewDN,$DNRequest.NewParentDistinguishedName,$DNRequest.NewName )

											#$logger.Debug("Move old:{0}|new:{1}",$DN,$NewDN)
											$result = $global:LdapConnection.SendRequest($DNRequest)
											if( $result.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success)
											{
												#$logger.error("Error ModifyDNRequest {0} `nDN: {1} `nNewDN: {2}",$result.ErrorMessage,$Entry.DistinguishedName,$NewDN)
												
												$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorConnectedDirectoryError)
												$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
												$Request = $null
												break
											}
											#Set new DN on request
											$Request.DistinguishedName = $NewDN
											$DN_ModifyRequest = $true
										}else{
											#Replace attibuts
											$DirectoryModification  = new-object System.DirectoryServices.Protocols.DirectoryAttributeModification
											$first = $true
											$DirectoryModification.Name = $AttributeName
											
											foreach($ValueChange in $CSEntry.AttributeChanges[$AttributeName].ValueChanges){
												#$logger.Debug("{0}:{1} | {2}",$AttributeName,$ValueChange.Value.ToString(),$ValueChange.ModificationType.ToString())
												switch($ValueChange.ModificationType.ToString()){
													"Delete" {
														if($first){
															$DirectoryModification.Operation = [System.DirectoryServices.Protocols.DirectoryAttributeOperation]::Delete
														}
													}
													"Add" {
															$first = $false
															$DirectoryModification.Operation = [System.DirectoryServices.Protocols.DirectoryAttributeOperation]::Replace
															[void]$DirectoryModification.Add($ValueChange.Value.ToString())
														}
												}
											}
											[void]$Request.Modifications.Add($DirectoryModification)
										}
									}
									if($Request.Modifications.Count -eq 0){
										$Request = $null
									}	
									break
								}
				#"Update" 		{
				#					#break
				#				}
				"Delete" 		{
									#Delete entry
									$Request = new-object System.DirectoryServices.Protocols.DeleteRequest($CSEntry.DN.ToString())
									break
								}
			}

			#Send request
			if($Request)
			{
				$result = $global:LdapConnection.SendRequest($Request)
				
				if( $result.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success)
				{
					$logger.error("{0} | {1}",$result.ErrorMessage,$Request.DistinguishedName)
					
					$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorConnectedDirectoryError)
					$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
				}
			}elseif(-NOT $DN_ModifyRequest)
			{
				#error no Request for entry
				$logger.error("No Request for entry DN: {0}",$DN)
					
				$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorConnectedDirectoryError)
				$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
			}
		}
	}
	Catch
	{
		$global:logger.Error($_.Exception.Message)
		$global:logger.Error($_.Exception.Source)
		$global:logger.Error($_.Exception.StackTrace)
		$global:logger.Error("LDAP-DIRX.ps1 ScriptLineNumber:{0}",$_.InvocationInfo.ScriptLineNumber)
		
		$CSEntryChangeResult = [Microsoft.MetadirectoryServices.CSEntryChangeResult]::Create($CSEntry.Identifier,$CSEntry.AttributeChanges,[Microsoft.MetadirectoryServices.MAExportError]::ExportErrorConnectedDirectoryError)
		$PutExportEntriesResults.CSEntryChangeResults.Add($CSEntryChangeResult)
	}
	$PutExportEntriesResults
}

Function IMAExtensible2CallExport.CloseExportConnection{
	Param
    (
		$CloseExportConnectionRunStep
	)
	try
	{
		#$global:logger.Debug("Run CloseExport")
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