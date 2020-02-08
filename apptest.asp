<%   
Dim vRun, vMsg

vMsg = ""
vRun = Trim(Request("run"))    'get form input

'If form value was submitted, run the task
If vRun = "Run Task" Then 
   RunJob 
End If  
 
'Routine to execute our task
Sub RunJob     
    Dim objTaskService, objRootFolder, objTask 

    'create instance of the scheduler service
    Set objTaskService = Server.CreateObject("Schedule.Service")    
    
    'connect to the service
    objTaskService.Connect            

    'go to our task folder, use just "\" if you saved it under the root folder
    Set objRootFolder = objTaskService.GetFolder("WebTasks")    
    'reference our task
    Set objTask = objRootFolder.GetTask("test")            

    'run it
    objTask.Run vbNull  
    vMsg = "Submitted"
    
    'clean up
    Set objTaskService = Nothing    
    Set objRootFolder = Nothing
    Set objTask = Nothing  
End Sub

%>
<html>
<body> 
<form method="post" action="apptest.asp"> 
  <p>Click to run the task: <input type="submit" value="Run Task" name="run" /></p> 
  <p>[<%= run %>]</p> 
  <p style="font-weight:bold; color:#006600;"><%= vMsg %></p>  
</form> 
</body>
</html>
