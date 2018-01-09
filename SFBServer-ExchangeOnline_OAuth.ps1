	import-module MSOnlineExtended

	#Export certificate to be used for OAuth
	$CertPath = "C:\Temp\SFBoauth.cer"
	$Thumbprint = (Get-CsCertificate -Type OAuthTokenIssuer).Thumbprint
	$OAuthCert = Get-ChildItem -Path Cert:\LocalMachine\My\$Thumbprint
	
	Export-Certificate -Cert $OAuthCert -FilePath $CertPath -Type CERT

	#Store certificate in appropriate format
	$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate
	$certificate.Import("C:\Temp\SFBauth.cer")
	$binaryValue = $certificate.GetRawCertData()
	$credentialsValue = [System.Convert]::ToBase64String($binaryValue)

	#Date variables for certificate start and expiration
	#Start can not be prior to today, and expriation cannot be greater than one year later than today, even if the certificate itself is valid beyond that date range
	$DateStart = Get-Date -format MM/dd/yyyy
	$DateExpire = (Get-Date).AddDays(364)
	$DateExpire = $DateExpire.ToString("MM/dd/yyyy")

	#Tenant ID for autodiscover URL
	$TenantID = Get-MSOLCompanyInformation | select objectID
	$TenantID = ($TenantID).ObjectID.GUID

	#If autodiscover config exists, grab the existing Autodiscover URL
	if ((Get-CSOauthConfiguration).ExchangeAutodiscoverURL)
	{
		$AutodiscoverURL = (Get-CSOauthConfiguration).ExchangeAutodiscoverURL
	}
	#If autodiscover config does not exist, prompt for the domain name and create the Office 365 default autodisover URL
	else 
	{
		$Domain = Read-Host "Enter the organization's verified domain name in Office 365" 
		$AutodiscoverURL = "http://autodisover." + $Domain + "/autodiscover/autodiscover.svc"   
	}

	#If allowed domains exist, grab the existing domain
	if ((Get-CSOauthConfiguration).ExchangeAutodiscoverAllowedDomains
	{
		$AutodiscoverAllowedDomains = (Get-CSOAuthConfiguration).ExchangeAutodiscoverAllowedDomains
	}
	#If allowed domains does not exist, add in the Office 365 default
	else 
	{
		$AutodiscoverAllowedDomains = "*.outlook.com"   
	}

	#If there is an existing OAuth server configured, remove it
	#OAuth server configuration must be done from scratch when certificate/credential is expired
	if ((Get-CsOauthServer).Identity -match "microsoft.sts") 
    {
		Write-Output "Removing existing OAuth Server"
		Remove-CsOauthServer -Identity microsoft.sts
	}

	#If there is an existing Partner Application, remove it
	#Partner application configuration must be done from scratch when certificate/credential is expired
	if ((Get-CsPartnerApplication).Identity -match "microsoft.exchange") 
    {
		Write-Output "Removing existing Partner Application"
		Remove-CsPartnerApplication -Identity microsoft.exchange
	}

	#Create the new Partner Application
	New-CsPartnerApplication -Identity microsoft.exchange -ApplicationIdentifier 00000002-0000-0ff1-ce00-000000000000 -ApplicationTrustLevel Full -UseOAuthServer

	#Create the new credentail using the certifcate data
	New-MsolServicePrincipalCredential -AppPrincipalId 00000004-0000-0ff1-ce00-000000000000 -Type Asymmetric -Usage Verify -Value $credentialsValue -StartDate $DateStart -EndDate $DateExpire

	#Create the new OAuth server using the Office 365 tenant ID
	New-CSOAuthServer microsoft.sts -MetadataURL "https://accounts.accesscontrol.windows.net/" + $TenantID + "/metadata/json/1"

	#Ensure OAuth server configuration is using the correct autodiscover Get-MSOLCompanyInformation
	#Autodiscover URL cannot use HTTPS
	Set-CSOAuthConfiguration -ExchangeAutodiscoverURL $ExchangeAutodiscoverURL -ExchangeAutodiscoverAllowedDomains $ExchangeAutodiscoverAllowedDomains
