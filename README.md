<div align="center">

## Outlook Contacts as ADO RecordSet


</div>

### Description

The purpose of this code is to read the Outlook Contacts Folder into an ADO RecordSet. Often you want to let your app use either the Outlook Contacts or your own. This will give you a jump start.
 
### More Info
 
ADO RecordSet


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Gillham](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-gillham.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-gillham-outlook-contacts-as-ado-recordset__1-31761/archive/master.zip)





### Source Code

```
Sub Outlook_Contacts()
 Dim ADOConn As ADODB.Connection
 Dim ADORS As ADODB.Recordset
 Dim strConn As String
 Set ADOConn = New ADODB.Connection
 Set ADORS = New ADODB.Recordset
 With ADOConn
  'Change the Connection String below
  ' to the correct settings
  .ConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;Exchange 4.0;MAPILEVEL=Outlook Address Book\;PROFILE=Outlook;TABLETYPE=1;DATABASE=c:\temp"
  .Open
 End With
 With ADORS
  Set .ActiveConnection = ADOConn
  .CursorType = adOpenStatic
  .LockType = adLockReadOnly
  .Open "Select * from [Contacts]"
  .MoveFirst
  'Test: just loop thru the first contact
  Dim i As Long
  For i = 0 To ADORS.Fields.Count - 1
   Debug.Print ADORS(i).Name + _
   vbTab + Format(ADORS(i).Value)
  Next i
  .Close
 End With
 Set ADORS = Nothing
 ADOConn.Close
 Set ADOConn = Nothing
End Sub
```

