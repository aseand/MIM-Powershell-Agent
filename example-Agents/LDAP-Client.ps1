
Add-Type -Assembly System.DirectoryServices.Protocols

Function CreateLDAPConnection{
	Param
    (
		$Server,
		$Port,
		$NetworkCredential=[System.Net.CredentialCache]::DefaultCredentials,
		$AuthType=[System.DirectoryServices.Protocols.AuthType]::Negotiate,
		[switch]$SSL
	)

	try
	{	
		#$global:logger.Debug("{0} {1} {2} {3}",$Server,$Port,$NetworkCredentia.UserName,$AuthType)
		$LdapDirectoryIdentifier = new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($Server,$Port)
		$LdapConnection = new-object System.DirectoryServices.Protocols.LdapConnection($LdapDirectoryIdentifier, $NetworkCredential, $AuthType)
		$LdapConnection.AutoBind = $True
		$LdapConnection.SessionOptions.ProtocolVersion = 3
		#$LdapConnection.Bind
		if($SSL)
		{
			$LdapConnection.SessionOptions.SecureSocketLayer = $True
			$LdapConnection.SessionOptions.VerifyServerCertificate = {$True}
		}
		return $LdapConnection
	}
	Catch
	{
		if($global:logger){
			$global:logger.Error($_.Exception.Message)
			$global:logger.Error($_.Exception.Source)
			$global:logger.Error($_.Exception.StackTrace)
			$global:logger.Error("LDAP-Client.ps1 {0}",$_.InvocationInfo.ScriptLineNumber)
		}else{
			write-host $_.Exception.Message
			write-host $_.Exception.Source
			write-host $_.Exception.StackTrace
			write-host "LDAP-Client.ps1 :" $_.InvocationInfo.ScriptLineNumber
		}
		$_
	}
}

Function BeginPagedSearch{
	Param
    (
		$LdapConnection,
		$baseDN,
		$filter,
		$attributs,
		$pageSize = 1000,
		[System.DirectoryServices.Protocols.SearchScope]$SearchScope = [System.DirectoryServices.Protocols.SearchScope]::Subtree
	)
	try
	{
		#$StartTime = get-Date
		[void]($SearchRequest = new-object System.DirectoryServices.Protocols.SearchRequest($baseDN, $filter, $SearchScope, $null))
		[void]($attributs | % {$SearchRequest.Attributes.Add($_)})
		[void]($PageResultRequestControl = new-object System.DirectoryServices.Protocols.PageResultRequestControl($pageSize))
		[void]($SearchRequest.Controls.Add($PageResultRequestControl))
		
		#$global:logger("LDAP add time {0}",((get-Date)-$StartTime).TotalMilliseconds)
		
		$resultSet = New-Object System.Collections.ArrayList
		#$resultSet = New-Object System.Collections.Generic.HashSet[System.DirectoryServices.Protocols.SearchResultEntry]
		do
		{
			#$global:logger.Debug($LdapConnection.SessionOptions.HostName)
			$searchResponse = [System.DirectoryServices.Protocols.SearchResponse] $LdapConnection.SendRequest($SearchRequest)
			$searchResponse.Controls | ? { $_ -is [System.DirectoryServices.Protocols.PageResultResponseControl] } | % {
				$PageResultRequestControl.Cookie = ([System.DirectoryServices.Protocols.PageResultResponseControl]$_).Cookie
			}
			#write-host "LDAP SearchRequest SendRequest done... " ((get-Date)-$StartTime).TotalSeconds
			$StartTime = get-Date
			if($searchResponse.Entries.count -gt 0)
			{
				$resultSet.AddRange($searchResponse.Entries)
			}
			#$global:logger.Debug("LDAP Response to list {0}",((get-Date)-$StartTime).Ticks)
		}
		while ($PageResultRequestControl.Cookie.length -gt 0)
		return ,$resultSet
	}
	Catch
	{
		if($global:logger){
			$global:logger.Error($_.Exception.Message)
			$global:logger.Error($_.Exception.Source)
			$global:logger.Error($_.Exception.StackTrace)
			$global:logger.Error("LDAP-Client.ps1 {0}",$_.InvocationInfo.ScriptLineNumber)
		}else{
			write-host $_.Exception.Message
			write-host $_.Exception.Source
			write-host $_.Exception.StackTrace
			write-host "LDAP-Client.ps1 :" $_.InvocationInfo.ScriptLineNumber
		}
		$_
	}
	#Finally
	#{
	#}
}

Function BeginPagedSearch-AsJob{
	Param
    (
		$LdapConnection,
		$baseDN,
		$filter,
		$attributs,
		$pageSize,
		$Job
	)
	try
	{
		[void]($SearchRequest = new-object System.DirectoryServices.Protocols.SearchRequest($baseDN, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree, $null))
		[void]($attributs | % {$SearchRequest.Attributes.Add($_)})
		[void]($PageResultRequestControl = new-object System.DirectoryServices.Protocols.PageResultRequestControl($pageSize))
		[void]($SearchRequest.Controls.Add($PageResultRequestControl))

		$resultSet= @()
		do
		{
			$searchResponse = [System.DirectoryServices.Protocols.SearchResponse] $LdapConnection.SendRequest($SearchRequest)
			$searchResponse.Controls | ? { $_ -is [System.DirectoryServices.Protocols.PageResultResponseControl] } | % {
				$PageResultRequestControl.Cookie = ([System.DirectoryServices.Protocols.PageResultResponseControl]$_).Cookie
			}
			if($searchResponse.Entries.count -gt 0)
			{
				[void](Start-Job -ScriptBlock $Job -Args $searchResponse)
			}
		}
		while ($PageResultRequestControl.Cookie.length -gt 0)
	}
	Catch
	{
		write-host $_.Exception.Message
		write-host $_.Exception.Source
		write-host $_.Exception.StackTrace
		write-host $_.InvocationInfo.ScriptLineNumber
	}
	Finally
	{
	}
}

Function BeginPagedSearchAsync{
	Param
    (
		$LdapConnection,
		$baseDN,
		$filter,
		$attributs,
		$pageSize
	)
	try
	{
		[void]($SearchRequest = new-object System.DirectoryServices.Protocols.SearchRequest($baseDN, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree, $null))
		[void]($attributs | % {$SearchRequest.Attributes.Add($_)})
		[void]($PageResultRequestControl = new-object System.DirectoryServices.Protocols.PageResultRequestControl($pageSize))
		[void]($SearchRequest.Controls.Add($PageResultRequestControl))

		$done = {

		
			write-host "Done" $_
			#[Action]{}|%{$_.BeginInvoke($null, $null)}
		}
		#
		$asyncCallback = [AsyncCallback]{
			Param ([IAsyncResult]$result)
			try
			{
				$searchResponse = $LdapConnection.EndSendRequest($result)
				#$pageAction($searchResponse)
				write-host "asyncCallback" $searchResponse.Entries.count
				$cookie = $null
				$searchResponse.Controls | ? { $_ -is [System.DirectoryServices.Protocols.PageResultResponseControl] } | % {
					$cookie = ([System.DirectoryServices.Protocols.PageResultResponseControl]$_).Cookie
				}

				if($cookie -ne $null -AND $cookie.length -ne 0)
				{
					$PageResultRequestControl.Cookie = $cookie
					$LdapConnection.BeginSendRequest($SearchRequest,[System.DirectoryServices.Protocols.PartialResultProcessing]::NoPartialResultSupport,$asyncCallback,$null)
				}
				else
				{
					#Invoke-Command -ScriptBlock $done -Args $null
				}
				
			}
			Catch
			{
				#Invoke-Command -ScriptBlock $done -Args $_
			}
		}
		write-host "BeginSendRequest"
		$LdapConnection.BeginSendRequest($SearchRequest,[System.DirectoryServices.Protocols.PartialResultProcessing]::NoPartialResultSupport,$asyncCallback,$null)
	}
	Catch
	{
		write-host $_.Exception.Message
		write-host $_.Exception.Source
		write-host $_.Exception.StackTrace
		write-host $_.InvocationInfo.ScriptLineNumber
	}
	Finally
	{
	}
}
