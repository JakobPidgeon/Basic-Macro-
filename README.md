# Basic-Macro-
# A basic macro code that was created for an assignment 

# The code when executed will open up Internet Explorer and play a Youtube video 

Sub OpenWebPage()
  
  Dim ie As Object
  
  Set ie = CreateObject("InternetExplorer.Application")
   
  ie.Visible = True
   
  Set WshShell = CreateObject("WScript.Shell")
  
  WshShell.Run "iexplore https://www.youtube.com/watch?v=dQw4w9WgXcQ"
 
  Msgbox "Credit Card Numbers  (1234567) (732738240 (124332) 

End Sub


