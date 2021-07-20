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

        $HighPart, $LowPart = [convert]::ToString( $Test, 16 ).PadLeft(16,'0') -split '(?=.{8}$)' | ForEach-Object { [int32]"0x$_" }

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

}