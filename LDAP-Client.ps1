
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
		$LdapDirectoryIdentifier = new-object System.DirectoryServices.Protocols.LdapDirectoryIdentifier($server,$port)
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
		write-host $_.Exception.Message
		write-host $_.Exception.Source
		write-host $_.Exception.StackTrace
		write-host $_.InvocationInfo.ScriptLineNumber
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
		
		#write-host "LDAP SearchRequest first " ((get-Date)-$StartTime).TotalSeconds
		
		$resultSet = New-Object System.Collections.ArrayList
		do
		{
			$searchResponse = [System.DirectoryServices.Protocols.SearchResponse] $LdapConnection.SendRequest($SearchRequest)
			$searchResponse.Controls | ? { $_ -is [System.DirectoryServices.Protocols.PageResultResponseControl] } | % {
				$PageResultRequestControl.Cookie = ([System.DirectoryServices.Protocols.PageResultResponseControl]$_).Cookie
			}
			#write-host "LDAP SearchRequest SendRequest done... " ((get-Date)-$StartTime).TotalSeconds
			if($searchResponse.Entries.count -gt 0)
			{
				$resultSet.AddRange($searchResponse.Entries)
			}
			#write-host "LDAP Response to list " ((get-Date)-$StartTime).TotalSeconds
		}
		while ($PageResultRequestControl.Cookie.length -gt 0)
		return ,$resultSet
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
