ScreenRes = funcScreenResolution() 
 
Function funcScreenResolution() 
    Set objIE = CreateObject("InternetExplorer.Application") 
    objIE.Navigate("about:blank") 
    Do Until objIE.readyState = 4: Loop 
        width = objIE.document.ParentWindow.screen.width 
        height = objIE.document.ParentWindow.screen.height 
    objIE.Quit 
    funcScreenResolution = array(width,height) 
End Function 
 
WScript.StdOut.WriteLine ScreenRes(0) & " x " & ScreenRes(1)
