<#
.SYNOPSIS
 Return Uninstall properties from the registry

.DESCRIPTION
 Return Uninstall properties from the registry

.PARAMETER Name
 The name to search for, supports wildcards

.EXAMPLE

 Get-UninstallEntry '*7-Zip*'

 Returns any entry where the DisplayName matches 7-Zip

#>
function Get-UninstallEntry {

    param(

        [Parameter(
            Mandatory = $true,
            Position = 1
        )]
        [string]
        $Name

    )
    
    Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like $Name }

}

<#
.SYNOPSYS
 Attempts to run an MSI uninstall with no output

.DESCRIPTION
 Attempts to run an MSI uninstall with no output

.PARAMETER QuietUninstallString
 The QuiteUninstallString from Get-UninstallEntry, will also fallback to UninstallString. Supports pipeline input.

.EXAMPLE

 Get-UninstallEntry '*7-Zip*' | Invoke-MsiQuietUninstall

#>
function Invoke-MsiQuietUninstall {

    [CmdletBinding()]
    param(
    
        [Parameter( Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true )]
        [Alias( 'UninstallString' )]
        [string]
        $QuietUninstallString
    
    )

    process {
    
        $FilePath, [string[]]$ArgumentList = __ExpandCommandLine $QuietUninstallString

        $Command = Get-Command $FilePath -ErrorAction SilentlyContinue

        if ( -not $Command -or $Command.Name -ne 'msiexec.exe' ) {

            Write-Warning "Uninstall command is not supported: $FilePath"
            return

        }

        [System.Collections.ArrayList]$SwitchParams = @()
        [System.Collections.ArrayList]$NamedParams = @()

        for ( $i = 0; $i -lt $ArgumentList.Count; $i ++ ) {

            # if argument is a bare /I or /X next param is always a path
            if ( $ArgumentList[$i] -eq '/I' -or $ArgumentList[$i] -eq '/X' ) {

                Write-Verbose "Replacing /I with /X"
                
                $SwitchParams.Add( '/X' ) > $null

                $i ++

                Write-Verbose "Adding path: $($ArgumentList[$i])"

                $SwitchParams.Add( $ArgumentList[$i] ) > $null

                continue
                
            }

            # if argument is a /I or /X followed by any string we just change to an /X
            if ( $ArgumentList[$i] -like '/[IX]*' ) {

                Write-Verbose "Replacing /I with /X"

                $SwitchParams.Add( $ArgumentList[$i].Replace( '/I', '/X' ) ) > $null

                continue

            }

            # if argument is a /L followed by any string we just include it
            if ( $ArgumentList[$i] -match '^/(log|l[iwearucmopvx+!\*]+)' ) {

                Write-Verbose "Including logging directive: $($ArgumentList[$i])"

                $SwitchParams.Add( $ArgumentList[$i] ) > $null

                # if next param does not start with / or have an "=" we append it right after the switch
                if ( $ArgumentList[$i+1][0] -ne '/' -and $ArgumentList[$i+1].IndexOf('=') -eq -1 ) {

                    $i ++

                    Write-Verbose "Adding path: $($ArgumentList[$i])"

                    $SwitchParams.Add( $ArgumentList[$i] ) > $null

                }

                continue

            }

            # if argument is a quiet switch we skip
            if ( $ArgumentList[$i] -eq '/quiet' -or $ArgumentList[$i] -eq '/passive' -or  $ArgumentList[$i] -like '/q[nbrf]*' ) {

                Write-Verbose "Skipping quiet directive: $($ArgumentList[$i])"

                continue

            }

            # if argument is a restart switch we skip
            if ( $ArgumentList[$i] -like '/[npf]*restart' ) {

                Write-Verbose "Skipping restart directive: $($ArgumentList[$i])"

                continue

            }

            # otherwise any other switches we add it
            if ( $ArgumentList[$i] -like '/*' ) {

                Write-Verbose "Adding unknown switch parameter: $($ArgumentList[$i])"

                $SwitchParams.Add( $ArgumentList[$i] ) > $null

                # if next param does not start with / or have an "=" we append it right after the switch
                if ( $ArgumentList[$i+1][0] -ne '/' -and $ArgumentList[$i+1].IndexOf('=') -eq -1 ) {

                    $i ++

                    Write-Verbose "Adding path: $($ArgumentList[$i])"

                    $SwitchParams.Add( $ArgumentList[$i] ) > $null

                }

                continue

            } else {

                Write-Verbose "Processing named param: $($ArgumentList[$i])"

                # in all other cases we add to $NamedParams
                $NamedParams.Add( $ArgumentList[$i] ) > $null

            }

        }

        Write-Verbose "Adding /qn /norestart"
        $SwitchParams.Add( '/qn' ) > $null
        $SwitchParams.Add( '/norestart' ) > $null

        [string[]]$ArgumentList = $SwitchParams + $NamedParams

        return __InvokeUninstallCommand -FilePath 'msiexec.exe' -ArgumentList $ArgumentList -Timeout $Timeout
        
    }

}


<#
.SYNOPSYS
 Attempts to run an EXE uninstall with no output

.DESCRIPTION
 Attempts to run an EXE uninstall with no output

.PARAMETER QuietUninstallString
 The QuiteUninstallString from Get-UninstallEntry. Supports pipeline input.

.PARAMETER UninstallString
 The UninstallString from Get-UninstallEntry. Fall back support for pipeline input if QuietUninstallString is not provided. Typically you also need to supply -SilentParams.

.PARAMETER SilentParams
 The parameters to supply to the uninstaller to make it silent.

.PARAMETER ReplaceParams
 Replace the existing parameters with the values in -SilentParams vs adding them.

.EXAMPLE

 Get-UninstallEntry '*7-Zip*' | Invoke-ExeQuietUninstall -SilentParams '/S'

#>
function Invoke-ExeQuiteUninstall {

    [CmdletBinding( DefaultParameterSetName = 'QuietUninstallString' )]
    param(
    
        [Parameter(
            ParameterSetName = 'QuietUninstallString',
            Mandatory = $true,
            Position = 1,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]
        $QuietUninstallString,

        [Parameter(
            ParameterSetName = 'UninstallString',
            Mandatory = $true,
            Position = 1,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]
        $UninstallString,

        [Parameter(
            ParameterSetName = 'UninstallString'
        )]
        [string[]]
        $SilentParams,

        [Parameter(
            ParameterSetName = 'UninstallString'
        )]
        [switch]
        $ReplaceParams,

        [int]
        $Timeout = 900
    
    )

    process {
    
        if ( $PSCmdlet.ParameterSetName -eq 'QuietUninstallString' ) {
        
            Write-Verbose "Attempting Quiet Uninstall: $QuietUninstallString"

            $FilePath, [string[]]$ArgumentList = __ExpandCommandLine $QuietUninstallString
        
        # regular uninstall
        } else {
        
            Write-Verbose "Attempting Uninstall: $UninstallString"

            $FilePath, [string[]]$ArgumentList = __ExpandCommandLine $UninstallString

            if ( $SilentParams.Count -gt 0 ) {

                if ( $ReplaceParams ) {

                    $ArgumentList = $SilentParams

                } else {

                    $ArgumentList = $ArgumentList + $SilentParams

                }

            }
        
        }

        return __InvokeUninstallCommand -FilePath $FilePath -ArgumentList $ArgumentList -Timeout $Timeout

    }

}

<#
.SYNOPSIS
 Retrieves properties from MSI installer file

.DESCRIPTION
 Retrieves properties from MSI installer file
#>
function Get-MsiFileProperties {
    
    [CmdletBinding()]
    param(

        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $true
        )]
        [System.IO.FileInfo[]]
        $Path

    )

    begin {

        $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer

        $Query = 'SELECT * FROM Property'

    }

    process {

        $Path | ForEach-Object {

            try {

                $Properties = [ordered]@{
                    Name = $_.Name
                    Path = $_.FullName
                }
                
                $MsiDatabase = $WindowsInstaller.GetType().InvokeMember( 'OpenDatabase', 'InvokeMethod', $null, $WindowsInstaller, @( $_.FullName, 0 ) )

                $OpenView = $MSIDatabase.GetType().InvokeMember( 'OpenView', 'InvokeMethod', $null, $MSIDatabase, $Query )

                $OpenView.GetType().InvokeMember( 'Execute', 'InvokeMethod', $null, $OpenView, $null )

                while ( $Record = $OpenView.GetType().InvokeMember( 'Fetch', 'InvokeMethod', $null, $OpenView, $null ) ) {

                    $Key   = $Record.GetType().InvokeMember( 'StringData', 'GetProperty', $null, $Record, 1 )
                    $Value = $Record.GetType().InvokeMember( 'StringData', 'GetProperty', $null, $Record, 2 )
        
                    $Properties[$Key] = $Value

                }

                $MSIDatabase.GetType().InvokeMember( 'Commit', 'InvokeMethod', $null, $MSIDatabase, $null )

                $OpenView.GetType().InvokeMember( 'Close', 'InvokeMethod', $null, $OpenView, $null )

                [pscustomobject]$Properties

            } catch {
            
                Write-Warning $_.Exception.Message
            
            }

        }

    }

    end {

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject( $WindowsInstaller ) > $null
        [System.GC]::Collect()

    }

}

function __ExpandCommandLine {

    param(

        [Parameter( Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true )]
        [string]
        $UninstallString

    )

    $UninstallString -replace '(?<!`)([\(\)\[\]{}@&$;])', '`$1' | ForEach-Object {

        Invoke-Expression "& {`$args} $_"

    }

}

function __InvokeUninstallCommand {

    param(

        [Parameter(
            Mandatory = $true,
            Position = 1
        )]
        [string]
        $FilePath,

        [string[]]
        $ArgumentList,
        
        [int]
        $Timeout = 900

    )

    $UninstallProcess = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -WindowStyle Hidden -PassThru

    try {

        $UninstallProcess | Wait-Process -Timeout $Timeout -ErrorAction SilentlyContinue -ErrorVariable UninstallTimeout

        if ( $UninstallTimeout ) {

            $UninstallProcess | Stop-Process -Force

            Write-Warning "Cancelled uninstall due to timeout after $Timeout seconds"

        }

    } finally {

        if ( -not $UninstallProcess.HasExited ) {

            $UninstallProcess | Stop-Process -Force

            Write-Warning "Cancelled uninstall due to user termination"

        }

    }

    return $UninstallProcess.ExitCode

}
