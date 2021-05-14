
#Add-Type -Path 'C:\Program Files\Microsoft Forefront Identity Manager\2010\Synchronization Service\UIShell\Microsoft.MetadirectoryServicesEx.dll'

Function Initialize{
	Param
    (
		$logger,
		$MAName
	)
	$global:logger = $logger
	$global:MAName = $MAName
	$global:logger.Info("Run Initialize")
}

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


