# Define the connection details for SQL Server
$server = "your_sql_server"           # SQL Server Name or IP
$database = "your_database"           # Database Name
$user = "your_username"               # SQL Server Username
$password = "your_password"           # SQL Server Password
$tableName = "your_table"             # Table Name where the data will be inserted

# File path to the pipe-delimited text file
$filePath = "C:\path\to\your\file.txt"

# Read the file into an array using pipe delimiter
$data = Import-Csv -Path $filePath -Delimiter '|'

# Create SQL connection string
$connectionString = "Server=$server;Database=$database;User Id=$user;Password=$password;"

# Loop through each row of the data and insert into SQL table
foreach ($row in $data) {
    $columns = $row.PSObject.Properties.Name -join ","
    $values = $row.PSObject.Properties.Value -join "','"
    $values = "'$values'"  # Enclose values in quotes for SQL query

    # Construct the SQL Insert query
    $query = "INSERT INTO $tableName ($columns) VALUES ($values);"

    # Execute the query
    try {
        Invoke-Sqlcmd -Query $query -ConnectionString $connectionString
        Write-Host "Row inserted successfully"
    } catch {
        Write-Host "Error inserting row: $_"
    }
}
