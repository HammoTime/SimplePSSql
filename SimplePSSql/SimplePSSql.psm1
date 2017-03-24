<#
    Version:        1.0.0.0
    Author:         Adam Hammond
    Creation Date:  25/03/2017
    Last Change:    Including update code.
    Description:    This library is provides simple functions for extracting, and
                    updating data from Microsoft SQL Server.
                    
    Link:           https://github.com/HammoTime/SimplePSSql
    License:        The MIT License (MIT)
#>

Function Select-SqlScalar
{
    <#
        .SYNOPSIS

        Gets a scalar (single) value from the specified database.

        .DESCRIPTION

        When given a query, returns the first value returned from the
        database.

        .PARAMETER Query

        The query to execute against the database.

        .PARAMETER ConnectionString

        The connection string used to connect to the server.

        .PARAMETER Parameters

        A hashtable of parameters. They key is the parameter key, and
        the value is the parameter value. Can include '@' although This
        isn't mandatory.

        .EXAMPLE

        A simple query with no parameters.

        $Query = 'SELECT TOP 1 [Value] FROM [MyDatabase].[dbo].[MyTable]'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        Select-SqlScalar $Query $ConnectionString

        .EXAMPLE

        A simple query with parameters.

        $Query = 'SELECT TOP 1 [Value] FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] = @Parameter'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = 'TestValue' }
        Select-SqlScalar $Query $ConnectionString $Parameters

        .EXAMPLE

        A simple query with a LIKE clause.

        $Query = 'SELECT TOP 1 [Value] FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] LIKE @Parameter'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = '%TestValue%' }
        Select-SqlScalar $Query $ConnectionString $Parameters

        .EXAMPLE

        A complex query, with multiple clauses.

        $Query = '
            SELECT TOP 1
                [Value]
            FROM
                [MyDatabase].[dbo].[MyTable]
            WHERE
                [MyColumn] LIKE @ParameterOne OR
                [MyColumn] = @ParameterTwo
        '
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@ParameterOne' = '%TestValue%'; '@ParameterTwo' = 'EqualsValue' }
        Select-SqlScalar $Query $ConnectionString $Parameters

    #>

    Param (
        [Parameter(Mandatory=$True)]
        [String]
        $Query,
        [Parameter(Mandatory=$True)]
        [String]
        $ConnectionString,
        [Hashtable]
        $Parameters = $null
    )

    $ReturnValue = $null

    # We don't catch exceptions, because we still want these directly exposed to the user.
    Try
    {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()

        $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)

        if($Parameters -ne $null)
        {
            ForEach($Key in $Parameters.Keys)
            {
                $ParameterName = ''

                if($Key.Substring(0, 1) -eq '@')
                {
                    $ParameterName = $Key
                } 
                else
                {
                    $ParameterName = '@' + $Key
                }

                $Command.Parameters.Add($ParameterName, $Parameters[$Key])
            }
        }

        $ReturnValue = $Command.ExecuteScalar()
    }
    Finally 
    {
        if($Connection -ne $null -and $Connection -is [System.IDisposable])
        {
            $Connection.Close()
        }

        if($Command -ne $null -and $Command -is [System.IDisposable])
        {
            $Command.Dispose()
        }
    }
    
    Return $ReturnValue
}

Function Select-SqlRows
{
    <#
        .SYNOPSIS

        Gets a DataTable based on a query from the specified database.

        .DESCRIPTION

        When given a query, returns a DataTable with the returned values
        from the database.

        .PARAMETER Query

        The query to execute against the database.

        .PARAMETER ConnectionString

        The connection string used to connect to the server.

        .PARAMETER Parameters

        A hashtable of parameters. They key is the parameter key, and
        the value is the parameter value. Can include '@' although This
        isn't mandatory.

        .EXAMPLE

        A simple query with no parameters.

        $Query = 'SELECT TOP 1000 * FROM [MyDatabase].[dbo].[MyTable]'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        Select-SqlRows $Query $ConnectionString

        .EXAMPLE

        A simple query with parameters.

        $Query = 'SELECT TOP 1000 * FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] = @Parameter'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = 'TestValue' }
        Select-SqlRows $Query $ConnectionString $Parameters

        .EXAMPLE

        A simple query with a LIKE clause.

        $Query = 'SELECT TOP 1000 * FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] LIKE @Parameter'
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = '%TestValue%' }
        Select-SqlRows $Query $ConnectionString $Parameters

        .EXAMPLE

        A complex query, with multiple clauses.

        $Query = '
            SELECT TOP 1000
                *
            FROM
                [MyDatabase].[dbo].[MyTable]
            WHERE
                [MyColumn] LIKE @ParameterOne OR
                [MyColumn] = @ParameterTwo
        '
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@ParameterOne' = '%TestValue%'; '@ParameterTwo' = 'EqualsValue' }
        Select-SqlRows $Query $ConnectionString $Parameters

    #>

    Param (
        [Parameter(Mandatory=$True)]
        [String]
        $Query,
        [Parameter(Mandatory=$True)]
        [String]
        $ConnectionString,
        [Hashtable]
        $Parameters = $null
    )

    $ReturnValue = New-Object System.Data.DataTable

    # We don't catch exceptions, because we still want these directly exposed to the user.
    Try
    {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()

        $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)

        if($Parameters -ne $null)
        {
            ForEach($Key in $Parameters.Keys)
            {
                $ParameterName = ''

                if($Key.Substring(0, 1) -eq '@')
                {
                    $ParameterName = $Key
                } 
                else
                {
                    $ParameterName = '@' + $Key
                }

                $Command.Parameters.Add($ParameterName, $Parameters[$Key])
            }
        }

        $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($Command)
        $DataAdapter.Fill($ReturnValue) | Out-Null
    }
    Finally 
    {
        if($Connection -ne $null -and $Connection -is [System.IDisposable])
        {
            $Connection.Close()
        }

        if($Command -ne $null -and $Command -is [System.IDisposable])
        {
            $Command.Dispose()
        }

        if($DataAdapter -ne $null -and $DataAdapter -is [System.IDisposable])
        {
            $DataAdapter.Dispose()
        }
    }
    
    Return $ReturnValue
}

Function Update-SqlTable
{
    <#
        .SYNOPSIS

        Updates a database using the given query.

        .DESCRIPTION

        When given a query, updates the database then returns the number
        of rows affected by the query.

        .PARAMETER Query

        The query to execute against the database.

        .PARAMETER ConnectionString

        The connection string used to connect to the server.

        .PARAMETER Parameters

        A hashtable of parameters. They key is the parameter key, and
        the value is the parameter value. Can include '@' although This
        isn't mandatory.

        .EXAMPLE

        A simple query with no parameters.

        $Query = "UPDATE [MyDatabase].[dbo].[MyTable] SET [MyColumn] = 'New Value'"
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $RowsAffected = Update-SqlTable $Query $ConnectionString

        .EXAMPLE

        A simple query with parameters.

        $Query = "UPDATE [MyDatabase].[dbo].[MyTable] SET [MyColumn] = 'New Value' WHERE [MyColumn] = @Parameter"
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = 'TestValue' }
        $RowsAffected = Update-SqlTable $Query $ConnectionString $Parameters

        .EXAMPLE

        Deleting a record.

        $Query = "DELETE FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] = @Parameter"
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $Parameters = { '@Parameter' = 'TestValue' }
        $RowsAffected = Update-SqlTable $Query $ConnectionString $Parameters

        .EXAMPLE

        Truncate a table.

        $Query = "TRUNCATE TABLE [MyDatabase].[dbo].[MyTable]"
        $ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
        $RowsAffected = Update-SqlTable $Query $ConnectionString

    #>

    $ReturnValue = 0

    Try
    {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()

        $Command = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)

        if($Parameters -ne $null)
        {
            ForEach($Key in $Parameters.Keys)
            {
                $ParameterName = ''

                if($Key.Substring(0, 1) -eq '@')
                {
                    $ParameterName = $Key
                } 
                else
                {
                    $ParameterName = '@' + $Key
                }

                $Command.Parameters.Add($ParameterName, $Parameters[$Key])
            }
        }

        $ReturnValue = $Command.ExecuteNonQuery()
    }
    Finally
    {
        if($Connection -ne $null -and $Connection -is [System.IDisposable])
        {
            $Connection.Close()
        }

        if($Command -ne $null -and $Command -is [System.IDisposable])
        {
            $Command.Dispose()
        }
    }

    Return $ReturnValue
}

Function Test-SqlConnection
{
    <# 
        .SYNOPSIS

        Tests to see if a server is accepting connections.

        .DESCRIPTION

        Attempts to connect to a server. If the connection fails,
        then false is returned, if it succeeds true is returned.

        .PARAMETER ConnectionString

        The connection string used to connect to the server.

        .EXAMPLE
        
        Test-SqlConnection 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
    #>

    Param (
        [Parameter(Mandatory=$True)]
        [String]
        $ConnectionString
    )

    $ReturnValue = $False

    Try
    {
        $Connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $Connection.Open()

        $ReturnValue = $True
    }
    Finally
    {
        if($Connection -ne $null -and $Connection -is [System.IDisposable])
        {
            $Connection.Close()
        }
    }

    Return $ReturnValue
}

<#
    Version:        1.1.0.0
    Author:         Adam Hammond
    Creation Date:  02/05/2016
    Last Change:    Created file.
    Description:    Contains utility functions that enhance SimplePSSql
                    
    Link:           https://github.com/HammoTime/SimplePSSql
    License:        The MIT License (MIT)
#>

Function Update-SimplePSSql
{
    <#
        .SYNOPSIS
        
        Updates the Simple PS Sql module from GitHub.
        
        .DESCRIPTION
        
        Checks the current releases on GitHub. If there is a new release, downloads the source
        files, copies them to the modules directory, and reloads the current runspace with
        the new module.
        
        .PARAMETER PreRelease
        
        If the latest release is a pre-release version and this switch is included, then
        it will install the pre-release version instead of the most recent production release.
         
        .LINK
         
        https://github.com/HammoTime/SimplePSSql/
    #>
    
    Param(
        [Switch]
        $PreRelease
    )
    Clear-Host
    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + '] [INFO]: Updating SimplePSSql.')

    $ModuleDirectory = $PSHome + '\Modules\SimplePSSql\'
    $TempDirectory = $Env:TEMP + '\SimplePSSql\'
    $ModuleZipLocation = $TempDirectory + 'SimplePSSql.zip'
    $ReleasesURL = 'https://git.io/vSTsX'

    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Retrieving release information from '$ReleasesURL'.")
    Try
    {
        $ReleasesPage = Invoke-WebRequest $ReleasesURL
    }
    Catch
    {
        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Couldn't retrieve release information (Invalid Request).") -ForegroundColor Red
    }

    if($ReleasesPage.StatusCode -eq 200)
    {
        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + '] [INFO]: Release information retrieved successfully.')
        $Releases = ConvertFrom-Json $ReleasesPage.Content
        if($PreRelease)
        {
            $LatestRelease = $Releases | Sort-Object -Property Published_At -Descending | Select-Object -First 1
        }
        else
        {
            $LatestRelease = $Releases | Sort-Object -Property Published_At -Descending | Where-Object { ([Boolean]$_.Prerelease) -eq $False } | Select-Object -First 1
        }
        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Current release version at $($LatestRelease.Tag_Name).")
        $InstalledVersion = (Get-Module -ListAvailable | Where-Object { $_.Name -eq 'SimplePSSql' }).Version
        $InstalledVersionNumber = 'v' + $InstalledVersion.Major + '.' + $InstalledVersion.Minor + '.' + $InstalledVersion.Build
        
        if($InstalledVersionNumber -eq $LatestRelease.Tag_Name)
        {
            Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Installed version at $($InstalledVersionNumber) (current).") -ForegroundColor Green
            Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: No update required.") -ForegroundColor Green
        }
        else
        {
            Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [WARN]: Installed version at $($InstalledVersionNumber) (out-of-date).") -ForegroundColor Yellow
            Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [WARN]: Update required.") -ForegroundColor Yellow
            Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Creating temporary directory at '$TempDirectory'.")
            New-Item $TempDirectory -ItemType Directory -Force | Out-Null
        
            Try
            {
                Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + '] [INFO]: Downloading release zip file.')
                Invoke-WebRequest $LatestRelease.ZipBall_Url -OutFile $ModuleZipLocation
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + '] [INFO]: Unpacking zip file.')
                [System.IO.Compression.ZipFile]::ExtractToDirectory($ModuleZipLocation, $TempDirectory)
                $ExtractedFiles = Get-ChildItem ($TempDirectory + (Get-ChildItem $TempDirectory -Directory | Select-Object -First 1).Name + '\SimplePSSql\')

                Try
                {
                    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Deleting existing module files at '$ModuleDirectory'.")
                    Remove-Item $ModuleDirectory -Force -Recurse
                    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Creating module directory at '$ModuleDirectory'.")
                    New-Item $ModuleDirectory -ItemType Directory -Force | Out-Null

                    ForEach($ExtractedFile in $ExtractedFiles)
                    {
                        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: Copying '$($ExtractedFile.Name)' to module directory.")
                        Copy-Item -Path $ExtractedFile.FullName -Destination ($ModuleDirectory + $ExtractedFile.Name) -Force
                    }

                    Try
                    {
                        Remove-Module 'SimplePSSql' -ErrorAction Stop -Force
                        Import-Module 'SimplePSSql' -ErrorAction Stop -Force
                        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: 'SimplePSSql' successfully reimported into current session.") -ForegroundColor Green
                        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [INFO]: 'SimplePSSql' updated successfully.") -ForegroundColor Green
                    }
                    Catch
                    {
                        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Error reloading module into current session.") -ForegroundColor Red
                        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Please reload PowerShell to use new module.") -ForegroundColor Red
                    }
                }
                Catch
                {
                    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Error copying new module files to module directory.") -ForegroundColor Red
                }

                Try
                {
                    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + '] [INFO]: Deleting temporary files.')
                    Remove-Item $TempDirectory -Force -Recurse
                }
                Catch
                {
                    Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Error deleting temporary directory.") -ForegroundColor Red
                }
            }
            Catch
            {
                Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Error processing zip file.") -ForegroundColor Red
            }
        }
    }
    else
    {
        Write-Host ('[' + (Get-Date).ToString('dd/MM/yyyy HH:mm:ss') + "] [ERRR]: Couldn't retrieve release information (HTTP Status Code: $($ReleasesPage.StatusCode)).") -ForegroundColor Red
    }
}



Export-ModuleMember -Function Select-SqlScalar
Export-ModuleMember -Function Select-SqlRows
Export-ModuleMember -Function Update-SqlTable
Export-ModuleMember -Function Test-SqlConnection
Export-ModuleMember -Function Update-SimplePSSql