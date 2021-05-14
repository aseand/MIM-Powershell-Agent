
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

