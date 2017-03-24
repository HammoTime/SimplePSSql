# SimplePSSql
A simple library for executing Sql queries against Microsoft SQL Server.

## Install
```Powershell
# Must be run as an administrator.

```

## Updating
```Powershell
# Must be run as an administrator.
Update-SimplePSSql
```

## Usage

### Test-SqlConnection
```Powershell

Test-SqlConnection 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'

```

### Select-SqlScalar
```Powershell

$Query = 'SELECT TOP 1 [Value] FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] = @Parameter'
$ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
$Parameters = { '@Parameter' = 'TestValue' }
Select-SqlScalar $Query $ConnectionString $Parameters

```

### Select-SqlRows
```Powershell

$Query = 'SELECT TOP 1000 * FROM [MyDatabase].[dbo].[MyTable] WHERE [MyColumn] = @Parameter'
$ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
$Parameters = { '@Parameter' = 'TestValue' }
Select-SqlRows $Query $ConnectionString $Parameters

```

### Update-SqlTable
```Powershell

$Query = "UPDATE [MyDatabase].[dbo].[MyTable] SET [MyColumn] = 'New Value' WHERE [MyColumn] = @Parameter"
$ConnectionString = 'Server=LOCALHOST\INSTANCE_01;Trusted_Connection=True;Database=MyDatabase'
$Parameters = { '@Parameter' = 'TestValue' }
$RowsAffected = Update-SqlTable $Query $ConnectionString $Parameters

```