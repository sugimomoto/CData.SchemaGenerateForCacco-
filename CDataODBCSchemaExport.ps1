[void][System.Reflection.Assembly]::LoadWithPartialName("System.Data")

$connectionsString = "DSN=CData D365Sales Source"
$odbcCon = New-Object System.Data.Odbc.OdbcConnection($connectionsString)
$odbcCon.Open();

$odbcCmd = New-Object System.Data.Odbc.OdbcCommand
$odbcCmd.Connection = $odbcCon

# コマンド実行（SELECT）
$odbcCmd.CommandText = "
SELECT 
  'mysql' dbms,
  t.SchemaName TABLE_SCHEMA,
  t.TableName TABLE_NAME,
  c.ColumnName COLUMN_NAME,
  c.Ordinal ORDINAL_POSITION,
  c.DataTypeName DATA_TYPE,
  c.Length CHARACTER_MAXIMUM_LENGTH,
  CASE
    WHEN k.IsKey THEN 'PRIMARY KEY'
    WHEN k.IsForeignKey THEN 'FOREIGN KEY'
    ELSE null
  END CONSTRAINT_TYPE,
  k.ReferencedSchemaName REFERENCED_TABLE_SCHEMA,
  k.ReferencedTableName REFERENCED_TABLE_NAME,
  k.ReferencedColumnName REFERENCED_COLUMN_NAME
  
  FROM sys_tables t 
  
  LEFT JOIN sys_tablecolumns c ON t.SchemaName=c.SchemaName AND t.TableName=c.TableName 
  LEFT JOIN sys_keycolumns k  ON c.SchemaName=k.SchemaName AND c.TableName=k.TableName AND c.ColumnName=k.ColumnName
WHERE t.TableName IN ('accounts','contacts','opportunities');
"


#SQLを実行し結果をDataSetやDataTableに格納
$odbcDataAdapter = New-Object -TypeName System.Data.Odbc.OdbcDataAdapter($odbcCmd)
$dataSet = New-Object -TypeName System.Data.DataSet
$odbcDataAdapter.Fill($dataSet) > $null      # 実行結果を破棄する
 
#データセットのデータテーブルをCSV出力する
$dataSet.Tables[0] | export-csv export_csvFile.csv -notypeinformation -Encoding Default

# コマンドオブジェクト破棄
$odbcCmd.Dispose()

# DB切断
$odbcCon.Close()
$odbcCon.Dispose()