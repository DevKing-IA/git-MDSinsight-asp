<%

'Usage

'Returning Data

'Set db = New dbaccess
'Call db.DBOpenAccess("/dblocation/dbfile.mdb")
'Call db.OpenRec() 
'arrArray = db.ExecuteSQL("SELECT * FROM [widgets];")
'Call db.CloseRec() 
'Call db.DBClose()
'Response.Write(arrArray(0,0))
'Erase arrArray

'-----------OR-------------

'Update Functions

'Set db = New dbaccess
'Call db.DBOpenAccess("/dblocation/dbfile.mdb")
'Call db.ExecuteUpdateSQL("DELETE FROM [widgets] WHERE [ID] = 1;")
'Call db.DBClose()

'++----------------------------------------------------------------
'++Class dbaccess
'++Author: Justin Owens
'++Date: 08/01/2002
'++You may use this class when developing your own applications 

'++as long as this section of comments remains with the class.
'++----------------------------------------------------------------

Class dbaccess
'Declarations
Private cnnObj
Private objRec
Private strConnStr

'Subs
'Class Initialization
Private Sub Class_Initialize()
     'Empty
End Sub

'Terminate Class
Private Sub Class_Terminate()
     'Empty
End Sub



'Open SQL Database
Public Sub DBOpenSQL()
     'strConnStr = InsightCnnString
     strConnStr =Session("ClientCnnString")
     Set cnnObj = server.CreateObject("ADODB.Connection")
     cnnObj.Open strConnStr
End Sub

'Close Database
Public Sub DBClose()
     cnnObj.Close
     Set cnnObj = Nothing
End Sub

'Functions
'Open Recordset
Public Function OpenRec()
     Set objRec = server.CreateObject("ADODB.Connection")
End Function

'Execute SQL
'Returns GetRows array - if no recset returned, it returns false.
Public Function ExecuteSQL(strSQLStatement)
'response.write(strSQLStatement)
     Set objRec = cnnObj.Execute(strSQLStatement)
     If Not objRec.EOF Then
          ExecuteSQL = objRec.GetRows() 
     Else
          ExecuteSQL = False
     End If
End Function

'Execute SQL
'Updates, inserts or deletes records in tables
Public Function ExecuteUpdateSQL(strSQLStatement)
     Set objRec = cnnObj.Execute(strSQLStatement)
     
End Function

Public Function ExecuteIdentity(strSQLStatement)
     Set objRec = cnnObj.Execute(strSQLStatement)

     ExecuteIdentity=objRec("IdentityInsert")
End Function

'Close RecordSet
Public Function CloseRec()
     objRec.close
     Set objRec = Nothing
End Function

End Class
%>