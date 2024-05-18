Attribute VB_Name = "MidpointPredToolCall"

Sub Button_CallMPT_Click()
                    'send job code and comp market to the text file
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set a = fs.CreateTextFile(Environ("AppData") & "\Microsoft\Addins\midpoint_prediction_params.txt", True)
                    a.writeline ("JobCode=" & JobCodes)
                    a.writeline ("CompMarket=" & compmarkets)
                    If ResultsCells <> "" Then
                        a.writeline ("ResultPlace=" & ResultsCells)
                        a.writeline ("Workbook=" & ActiveWorkbook.Name)
                        a.writeline ("Worksheet=" & ActiveSheet.Name)
                    End If
                    a.Close
                    
                    'check if critical folders already exist
                    Dim folderPath As String
                    folderPath = Environ("AppData") & "\Microsoft\Addins\Reports"
                    If Not FolderExists(folderPath) Then
                        ' Create the folder
                        MkDir folderPath
                    End If
                    folderPath = Environ("AppData") & "\Microsoft\Addins\Output Files"
                    If Not FolderExists(folderPath) Then
                        ' Create the folder
                        MkDir folderPath
                    End If
                    
                    'check if the python script already exists, if not, copy it from the server
                    scriptFileName = Environ("AppData") & "\Microsoft\Addins\MidpointPredictionTool-RunAlone.exe"
                    If Dir(scriptFileName) = "" Then
                        Midpoint_Prediction.CopyFileFromSharePoint
                    End If
                    
                    'check if the python script file is outdated, if it is, copy the latest version from the server
                    If FileDateTime(scriptFileName) <> FileDateTime("C:\Users\" & Split(Application.UserName, ", ")(1) & "_" & Split(Application.UserName, ", ")(0) & "\OneDrive - Company\MIT tools - Excel Add-in\Resources\MidpointPredictionTool-RunAlone.exe") Then
                        CopyFileFromSharePoint
                    End If
                    
                    'check if the reports are up-to-date, update them if they are not
                    updateReportsCall
                    
                    ' run python and the python script

                    mycommand = Chr(34) & scriptFileName & Chr(34)
                    Application.StatusBar = "Python is predicting your midpoint"
                    'Debug.Print (mycommand)
                    Shell (mycommand)


End Sub