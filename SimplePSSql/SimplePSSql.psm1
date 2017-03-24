<#
    Version:        1.0.0.0
    Author:         Adam Hammond
    Creation Date:  25/03/2017
    Last Change:    Adding to GitHub
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

Export-ModuleMember -Function Select-SqlScalar
Export-ModuleMember -Function Select-SqlRows
Export-ModuleMember -Function Update-SqlTable
Export-ModuleMember -Function Test-SqlConnection