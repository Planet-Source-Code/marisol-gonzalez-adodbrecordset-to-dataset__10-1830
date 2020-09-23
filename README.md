<div align="center">

## ADODBRecordSet To DataSet


</div>

### Description

Transform an ADODB.RecordSet to DataSet.
 
### More Info
 
1. Create a New Project Windows (Select a Windows Application)

2. You need a DataGrid ,Imports System.Data.OleDb and Imports ADODB.

3. Copy and paste code at form.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[marisol gonzalez](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marisol-gonzalez.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB\.NET, ASP\.NET
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__10-33.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marisol-gonzalez-adodbrecordset-to-dataset__10-1830/archive/master.zip)





### Source Code

```
Imports System.Data.OleDb
Imports ADODB
Public Class Form1
 Inherits System.Windows.Forms.Form
 Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  RecordSetToDataSet("ser_amp_bd2", "pubs")
 End Sub
 Private Function RecordSetToDataSet(ByVal Server As String, _
          ByVal DataBase As String)
  Dim cnn As Connection = New ADODB.Connection()
  Dim rs As Recordset = New ADODB.Recordset()
  '--Open connection--
  cnn.open("PROVIDER=SQLOLEDB;DATA SOURCE=" & Server & ";" & _
     "INITIAL CATALOG=" & DataBase & ";INTEGRATED SECURITY=SSPI;")
  '--create recordset--
  '--<any table> name of table in your database
  Dim sql As String = "select * from authors"
  rs.open(sql, cnn, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly, 0)
  Dim dadapter As New OleDbDataAdapter()
  Dim ddataset As New DataSet()
  '--Move rs to dataset--
  dadapter.Fill(ddataset, rs, "result")
  '--Fill a grid--
  grd.DataSource = ddataset.Tables("result")
  '--Close Connection--
  cnn.Close()
 End Function
End Class
```

