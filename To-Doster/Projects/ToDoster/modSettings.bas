
Attribute VB_Name = "modSettings"
#Const modSettings = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module
Public Type SettingType
    lLocation As Integer
    lProduct As Integer
    Location As String
    Products As String
    winTop As Integer
    winLeft As Integer
    winHeight As Integer
    winWidth As Integer
    col1Width As Integer
    col2Width As Integer
    col3Width As Integer
    col4Width As Integer
    col5Width As Integer
End Type

Public Settings As SettingType
    
Public colProducts As New Collection

Public Function IsOnList(tList, Item) As Integer

    Dim cnt As Integer
    Dim ItemFound As Integer
    cnt = 0
    ItemFound = -1
    Do Until cnt = tList.ListCount Or ItemFound > -1
        If LCase(Trim(tList.List(cnt))) = LCase(Trim(Item)) Then ItemFound = cnt
        cnt = cnt + 1
        Loop
    IsOnList = ItemFound

End Function
Public Function ClearCollection(ByRef nCol)
    Do Until nCol.Count = 0
        nCol.Remove 1
    Loop
End Function


Public Function LoadSettings() As String

    Dim nProduct As Variant
    
    If PathExists(AppPath & "todoster.bin", True) Then

        Dim num As Long
        num = FreeFile

        Open AppPath & "todoster.bin" For Binary Lock Read As #num

        Get #num, 1, Settings.lLocation
        Get #num, 3, Settings.lProduct

        Settings.Location = String(Settings.lLocation, Chr(0))
        Settings.Products = String(Settings.lProduct, Chr(0))

        Get #num, 1, Settings

        Settings.Location = Replace(Settings.Location, Chr(0), "")

        ClearCollection colProducts

        nProduct = Settings.Products
        Do While Len(nProduct) > 0
            colProducts.Add RemoveNextArg(nProduct, Chr(0))
        Loop

        Close #num

    End If
    
    If colProducts.Count = 0 Then colProducts.Add "Unknown"
    
    If Settings.Location = "" Then Settings.Location = AppPath & "\todoster.mdb"

End Function
    
Public Function SaveSettings()

    Dim nProduct As Variant

    If PathExists(AppPath & "todoster.bin", True) Then
        Kill AppPath & "todoster.bin"
    End If
    
    Settings.lLocation = Len(Settings.Location)
    Settings.Products = ""
    For Each nProduct In colProducts
        Settings.Products = Settings.Products & Chr(0) & nProduct
    Next
    
    If Left(Settings.Products, 1) = Chr(0) Then Settings.Products = Mid(Settings.Products, 2)
            
    Settings.lProduct = Len(Settings.Products)
    
    Dim num As Long
    num = FreeFile
    
    Open AppPath & "todoster.bin" For Binary Lock Write As #num
        Put #num, 1, Settings
    Close #num
    
End Function


'##########################################################################################################################
'##########################################################################################################################

Public Function IsURL(ByVal URL As String) As Boolean
    IsURL = Left(Trim(LCase(URL)), 4) = "http" Or Left(Trim(LCase(URL)), 5) = "https"
End Function

Public Function IsSSL(ByVal URL As String) As Boolean
    IsSSL = Left(Trim(LCase(URL)), 5) = "https"
End Function

Public Function GetServerName(ByVal URL As String) As String
    If IsURL(URL) Then
        URL = Mid(URL, 8)
        If Left(URL, 1) = "/" Then URL = Mid(URL, 2)
        If InStr(URL, "/") > 0 Then
            URL = Left(URL, InStr(URL, "/") - 1)
        End If
        If InStr(URL, "@") > 0 Then
            URL = Mid(URL, InStr(URL, "@") + 1)
        End If
        GetServerName = URL
    Else
        GetServerName = URL
    End If
End Function

Public Function GetServerWebForm(ByVal URL As String) As String
    If IsURL(URL) Then
        URL = Mid(URL, 8)
        If Left(URL, 1) = "/" Then URL = Mid(URL, 2)
        If InStr(URL, "/") > 0 Then
            URL = Mid(URL, InStr(URL, "/"))
        End If
        GetServerWebForm = URL
    Else
        GetServerWebForm = URL
    End If
End Function

Public Function GetUsername(ByVal URL As String) As String
    If IsURL(URL) Then
        URL = Mid(URL, 8)
        If Left(URL, 1) = "/" Then URL = Mid(URL, 2)
        If InStr(URL, "@") > 0 Then
            URL = Left(URL, InStr(URL, "@") - 1)
            If InStr(URL, ":") > 0 Then
                URL = Left(URL, InStr(URL, ":") - 1)
            End If
        End If
        GetUsername = URL
    Else
        GetUsername = URL
    End If
End Function
Public Function GetPassword(ByVal URL As String) As String
    If IsURL(URL) Then
        URL = Mid(URL, 8)
        If Left(URL, 1) = "/" Then URL = Mid(URL, 2)
        If InStr(URL, "@") > 0 Then
            URL = Left(URL, InStr(URL, "@") - 1)
            If InStr(URL, ":") > 0 Then
                URL = Mid(URL, InStr(URL, ":") + 1)
            End If
        End If
        GetPassword = URL
    Else
        GetPassword = URL
    End If
End Function

'##########################################################################################################################
'##########################################################################################################################
'ServerExecute posts ParamData (form) to ServerWebForm (file) on ServerName
'Returns the XML object the server responds with
Public Function ServerExecute(ByVal ServerName As String, ByVal ServerWebForm As String, ByVal ParamData As String, ByVal Username As String, ByVal Password As String, ByVal SSL As Boolean) As Variant
On Error GoTo catch
       
    Dim xml, postData
    Set xml = CreateObject("msxml.DOMDocument")
    
    xml.async = "false"
    xml.loadXML modInternet.PostToWebsite(ServerName, ServerWebForm, ParamData, Username, Password, , SSL)
    
    Set ServerExecute = xml
    
    Set xml = Nothing
    
    Exit Function
catch:
    Err.Clear
    ServerExecute = 0
    MsgBox "Unable to query web database, please ensure you have the correct access." & vbCrLf & "Example: https://username:password@www.server.com/todoster.asp", vbInformation
End Function
    
'PostExecute same as ServerExecute except it only returns a string of what the webpage response.write's
Public Function PostExecute(ByVal ServerName As String, ByVal ServerWebForm As String, ByVal ParamData As String, ByVal Username As String, ByVal Password As String, ByVal SSL As Boolean) As Variant
    On Error GoTo catch
    
    PostExecute = modInternet.PostToWebsite(ServerName, ServerWebForm, ParamData, Username, Password, , SSL)

    If IsNumeric(PostExecute) Then PostExecute = CLng(PostExecute)
    If PostExecute = "True" Or PostExecute = "False" Then PostExecute = CBool(PostExecute)
    
    Exit Function
catch:
    Err.Clear
    PostExecute = 0
    MsgBox "Unable to query web database, please ensure you have the correct access." & vbCrLf & "Example: https://username:passowrd@www.server.com/todoster.asp", vbInformation
End Function
    

     