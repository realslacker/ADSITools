using namespace System.DirectoryServices
using namespace System.DirectoryServices.AccountManagement
using namespace System.Security.Principal

# enum for username translation context
enum TranslateContext {
    Domain                  = 1  #ADS_NAME_INITTYPE_DOMAIN
    Server                  = 2  #ADS_NAME_INITTYPE_SERVER
    GlobalCatalog           = 3  #ADS_NAME_INITTYPE_GC
}

# enum for username translation types
enum TranslateType {
    DistinguishedName       = 1  #ADS_NAME_TYPE_1779
    CanonicalName           = 2  #ADS_NAME_TYPE_CANONICAL
    NTAccount               = 3  #ADS_NAME_TYPE_NT4
    DisplayName             = 4  #ADS_NAME_TYPE_DISPLAY
    DomainSimple            = 5  #ADS_NAME_TYPE_DOMAIN_SIMPLE
    EnterpriseSimple        = 6  #ADS_NAME_TYPE_ENTERPRISE_SIMPLE
    GUID                    = 7  #ADS_NAME_TYPE_GUID
    Unknown                 = 8  #ADS_NAME_TYPE_UNKNOWN
    UserPrincipalName       = 9  #ADS_NAME_TYPE_USER_PRINCIPAL_NAME
    CanonicalEx             = 10 #ADS_NAME_TYPE_CANONICAL_EX
    ServicePrincipalName    = 11 #ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME
    SID                     = 12 #ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME
}

Add-Type -AssemblyName System.DirectoryServices.AccountManagement

<#
.SYNOPSIS
Convert int64 or DateTime object to IADSLargeInteger

.DESCRIPTION
Convert int64 or DateTime object to IADSLargeInteger for updating AD objects using LDAP/ADSI

.PARAMETER Int64
Timestamp as an integer value

.PARAMETER DateTime
Timestamp as a datetime value
#>
function ConvertTo-IADsLargeInteger {
    
    param(

        [Parameter( Mandatory, ValueFromPipeline, ParameterSetName = 'FromInt64' )]
        [int64]
        $Int64,

        [Parameter( Mandatory, ValueFromPipeline, ParameterSetName = 'FromDateTime' )]
        [datetime]
        $DateTime

    )

    process {

        if ( $DateTime ) {
        
            $Int64 = $DateTime.ToFileTimeUtc()

        }

        $HighPart, $LowPart = [convert]::ToString( $Int64, 16 ).PadLeft(16,'0') -split '(?=.{8}$)' | ForEach-Object { [int32]"0x$_" }

        $ADsLargeInteger = New-Object -ComObject LargeInteger

        [void]$ADsLargeInteger.GetType().InvokeMember( 'HighPart', 'SetProperty', $null, $ADsLargeInteger, $HighPart )
        [void]$ADsLargeInteger.GetType().InvokeMember( 'LowPart',  'SetProperty', $null, $ADsLargeInteger, $LowPart  )

        return $ADsLargeInteger

    }

}


<#
.SYNOPSIS
Convert from and IADSLargeInteger to either an integer or DateTime object

.DESCRIPTION
Convert from and IADSLargeInteger to either an integer or DateTime object. IADSLargeIntegers are used by AD LDAP/ADSI interface for timestamps.

.PARAMETER InputObject
IADSLargeInteger ComObject

.PARAMETER OutputType
The selected output type
#>
function ConvertFrom-IADsLargeInteger {
    
    param(

        [Parameter( Mandatory, ValueFromPipeline )]
        [object]
        $InputObject,

        [ValidateSet( 'DateTime', 'DateTimeUtc', 'Int64', 'Long' )]
        [string]
        $OutputType = 'DateTime'

    )

    begin {

        $ADSI = [adsi]::new()

    }

    process {

        $Int64 = $ADSI.ConvertLargeIntegerToInt64( $InputObject )

        switch ( $OutputType ) {
            'DateTime'    { [datetime]::FromFileTimeUtc( $Int64 ).ToLocalTime() }
            'DateTimeUtc' { [datetime]::FromFileTimeUtc( $Int64 ) }
            'Long'        { [long]$Int64 }
            'Int64'       { $Int64 }
        }

    }

    end {

        $ADSI.Dispose()

    }

}

<#
.SYNOPSIS
Validate a credential.

.DESCRIPTION
Validate a credential in either the domain or machine context.

.PARAMETER Credential
The credential to test.

.PARAMETER MachineContext
Switch to indicate that the credential should be validated in the Machine
context.

.PARAMETER ContextOptions
A combination of one or more ContextOptions enumeration values the options
used to bind to the server. This parameter can only specify Simple bind with
or without SSL, or Negotiate bind.
#>
function Test-Credential {
    
    [CmdletBinding( DefaultParameterSetName = 'DomainContext' )]
    [OutputType( [bool] )]
    Param(
    
        [Parameter( Mandatory=$true )]
        [pscredential]
        $Credential,

        [Parameter( ParameterSetName = 'MachineContext' )]
        [switch]
        $MachineContext,

        [ContextOptions[]]
        $ContextOptions = 'Negotiate, Sealing'

    )

    Write-Verbose "UserName is $($Credential.UserName)"

    $Context = ( 'Domain', 'Machine' )[ $MachineContext.IsPresent ]

    Write-Verbose "Context is $Context"

    Write-Verbose "Context Options are $ContextOptions"

    $AuthObj = [PrincipalContext]::new( $Context )

    $AuthObj.ValidateCredentials(
        $Credential.UserName,
        $Credential.GetNetworkCredential().Password,    
        $ContextOptions
    )

    $AuthObj.Dispose()

}

<#
.SYNOPSIS
Converts user names between formats.

.DESCRIPTION
Converts user names between formats. Uses the ComObject NameTranslate.

.PARAMETER UserName
The user(s) to convert.

.PARAMETER InputType

.PARAMETER OutputType

.PARAMETER Credential
Credential used for binding to domain.
#>
function Convert-UserNameFormat {

    [CmdletBinding()]
    [OutputType( [string] )]
    param (
        
        [Parameter(Mandatory, Position=0)]
        [string[]]
        $UserName,

        [ValidateSet(
            'Unknown',
            'DistinguishedName',
            'CanonicalName',
            'NTAccount',
            'DisplayName',
            'DomainSimple',
            'EnterpriseSimple',
            'GUID',
            'UserPrincipalName',
            'CanonicalEx',
            'ServicePrincipalName',
            'SID'
        )]
        [TranslateType]
        $InputType = 'Unknown',

        [Parameter(Mandatory)]
        [ValidateSet(
            'DistinguishedName',
            'CanonicalName',
            'NTAccount',
            'DisplayName',
            'DomainSimple',
            'EnterpriseSimple',
            'GUID',
            'UserPrincipalName',
            'CanonicalEx',
            'ServicePrincipalName',
            'SID'
        )]
        [TranslateType]
        $OutputType,

        [TranslateContext]
        $Context = 'GlobalCatalog',

        [pscredential]
        $Credential

    )

    begin {

        $NameTranslateComObject = New-Object -ComObject 'NameTranslate'
        $NameTranslateType = $NameTranslateComObject.GetType()
    
        # if a credential is supplied we use the InitEx method
        if ( $Credential ) {

            # translate possible UPN to NT Account
            $Domain, $UserName = ([NTAccount]$UserName).
                Translate( [SecurityIdentifier] ).
                Translate( [NTAccount] ).Value.Split('\')

            $NameTranslateType.InvokeMember( 'InitEx', 'InvokeMethod', $null, $NameTranslateComObject, ( $Context, $null, $UserName, $Domain, $Credential.GetNetworkCredential().Password ) ) > $null

        # otherwise just init with the default user context
        } else {

            $NameTranslateType.InvokeMember( 'Init', 'InvokeMethod', $null, $NameTranslateComObject, ( $Context, $null ) ) > $null

        }

    }

    process {

        $UserName | ForEach-Object {

            # set the current user name
            $NameTranslateType.InvokeMember( 'Set', 'InvokeMethod', $null, $NameTranslateComObject, ( $InputType, [string]$_ ) ) > $null

            # if output type is SID we have to do extra conversion
            if ( $OutputType -eq [TranslateType]::SID ) {

                ([NTAccount]$NameTranslateType.InvokeMember( 'Get', 'InvokeMethod', $null, $NameTranslateComObject, [TranslateType]::NTAccount )).
                    Translate( [SecurityIdentifier] )

            # get the requested format
            } else {
    
                $NameTranslateType.InvokeMember( 'Get', 'InvokeMethod', $null, $NameTranslateComObject, $OutputType )

            }

        }

    }

}


<#
.SYNOPSIS
Return a DirectoryEntry object

.DESCRIPTION
Return a DirectoryEntry object

.PARAMETER DistinguishedName
Get directory entry with specific DistinguishedName

.PARAMETER ObjectSid
Get directory entry with specific ObjectSid

.PARAMETER ObjectGuid
Get directory entry with specific ObjectGuid

.PARAMETER RootDSE
Get RootDSE directory entry

.PARAMETER Server

.PARAMETER Credential
#>
function Get-ADSIDirectoryEntry {

    [CmdletBinding(
        PositionalBinding = $false
    )]
    [OutputType( [adsi] )]
    param (

        [Parameter(
            ParameterSetName = 'DistinguishedName',
            Mandatory
        )]
        [ValidateNotNullOrEmpty()]
        [Alias( 'DN' )]
        [string]
        $DistinguishedName,

        [Parameter(
            ParameterSetName = 'ObjectSid',
            Mandatory
        )]
        [ValidateNotNullOrEmpty()]
        [SecurityIdentifier]
        [Alias( 'SID', 'SecurityIdentifier' )]
        $ObjectSid,

        [Parameter(
            ParameterSetName = 'ObjectGuid',
            Mandatory
        )]
        [ValidateNotNullOrEmpty()]
        [Alias( 'GUID' )]
        [guid]
        $ObjectGuid,

        [Parameter(
            ParameterSetName = 'RootDSE',
            Mandatory
        )]
        [ValidateNotNullOrEmpty()]
        [switch]
        $RootDSE,

        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [switch]
        $GlobalCatalog,

        [pscredential]
        $Credential,

        [AuthenticationTypes[]]
        $AuthenticationTypes = 'Secure, Sealing',

        [Parameter( ValueFromRemainingArguments, DontShow )]
        $IgnoredParams

    )

    $DistinguishedName = switch ( $PSCmdlet.ParameterSetName ) {
        'DistinguishedName' { $DistinguishedName }
        'ObjectSid'         { '<SID={0}>' -f $ObjectSid }
        'ObjectGuid'        { '<GUID={0}>' -f $ObjectGuid }
        'RootDSE'           { 'RootDSE' }
    }

    $ConnectionString = ( $GlobalCatalog ? 'GC://' : 'LDAP://' ) + ( $Server ? "$Server/" : '' ) + $DistinguishedName

    try {

        if ( $Credential ) {
            [adsi]::new( $ConnectionString, $Credential.UserName, $Credential.GetNetworkCredential().Password, $AuthenticationTypes )
        } else {
            [adsi]$ConnectionString
        }

    } catch {

        Write-Error $_.Exception.Message

    }

}


<#
.SYNOPSIS
Return the specified naming context from the domain

.DESCRIPTION
Return the specified naming context from the domain

.PARAMETER NamingContext
Which naming context to return.

.PARAMETER DomainName
Which domain or server to query.

.PARAMETER Credential
#>
function Get-ADSINamingContext {

    [CmdletBinding()]
    [OutputType( [string] )]
    param (

        [ValidateSet( 'DefaultNamingContext', 'ConfigurationNamingContext', 'SchemaNamingContext' )]
        [ValidateNotNullOrEmpty()]
        [string]
        $NamingContext = 'DefaultNamingContext',

        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [switch]
        $GlobalCatalog,

        [pscredential]
        $Credential,

        [Parameter( ValueFromRemainingArguments, DontShow )]
        $IgnoredParams

    )

    return (Get-ADSIDirectoryEntry -RootDSE @PSBoundParameters).Get( $NamingContext )

}


<#


.PARAMETER AuthenticationTypes
The AuthenticationTypes enumeration specifies the types of authentication used in System.DirectoryServices.
#>
function Get-ADSISearcher {

    [CmdletBinding()]
    [OutputType( [adsisearcher] )]
    param (

        [Alias( 'Root' )]
        [AllowNull()]
        [string]
        $SearchRoot,

        [Alias( 'Scope' )]
        [SearchScope]
        $SearchScope = 'Subtree',

        [ValidateNotNullOrEmpty()]
        [Alias( 'Filter' )]
        [string]
        $SearchFilter = '(objectClass=*)',

        [string[]]
        $PropertiesToLoad,

        [ValidateNotNullOrEmpty()]
        [string]
        $Server,

        [switch]
        $GlobalCatalog,

        [pscredential]
        $Credential,

        [AuthenticationTypes[]]
        $AuthenticationTypes = 'Secure, Sealing'

    )

    if ( [string]::IsNullOrEmpty($SearchRoot) ) {
        $SearchRoot = Get-ADSINamingContext @PSBoundParameters
    }

    $DirectoryEntry = Get-ADSIDirectoryEntry -DistinguishedName $SearchRoot @PSBoundParameters

    [adsisearcher]::new(
        $DirectoryEntry,
        $SearchFilter,
        $PropertiesToLoad,
        $SearchScope
    )
    
}