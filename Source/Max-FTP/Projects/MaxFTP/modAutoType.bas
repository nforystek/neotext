Attribute VB_Name = "modAutoType"





#Const modAutoType = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Sub SetAutoTypeList(ByVal frm As Form, ByRef TypeCombo)

        Dim dbConn As New clsDBConnection
        Dim rs As New ADODB.Recordset
    
        dbConn.rsQuery rs, "SELECT * FROM History WHERE ParentID=" & dbSettings.CurrentUserID & " ORDER BY Stamp DESC;"
        
        TypeCombo.Clear

        Do Until rsEnd(rs)
                
            TypeCombo.AddItem rs("URL")

            rs.MoveNext
        Loop
    
        rsClose rs
        Set dbConn = Nothing
    
End Sub

Public Sub AddAutoTypeURL(ByVal URL As String)
        
    If (Not IsActiveAppFolder(URL)) Then
        
        Dim nUrl As New NTAdvFTP61.URL
        
        If (((Right(URL, 1) = "\") Or ((Right(URL, 1) = "/")) And (Not (nUrl.GetFolder(URL) = "/")))) And (Len(nUrl.GetFolder(URL)) > 3) Then URL = Left(URL, Len(URL) - 1)

        Set nUrl = Nothing
        
        Dim dbConn As New clsDBConnection
        Dim rs As New ADODB.Recordset
    
        dbConn.rsQuery rs, "SELECT * FROM History WHERE ParentID=" & dbSettings.CurrentUserID & " AND URL='" & Replace(URL, "'", "''") & "';"
        
        If rsEnd(rs) Then
            
            dbConn.rsQuery rs, "SELECT Count(*) as Cnt FROM History WHERE ParentID=" & dbSettings.CurrentUserID & ";"
        
            If rs("Cnt") >= dbSettings.GetProfileSetting("HistorySize") Then
                dbConn.rsQuery rs, "SELECT * FROM History WHERE ParentID=" & dbSettings.CurrentUserID & " ORDER BY Stamp;"
                
                dbConn.rsQuery rs, "DELETE FROM History WHERE ID=" & rs("ID") & ";"
            
            End If
            
            dbConn.rsQuery rs, "INSERT INTO History (ParentID, URL, Stamp) VALUES (" & dbSettings.CurrentUserID & ", '" & Replace(URL, "'", "''") & "', '" & Now & "');"
        
        Else
            dbConn.rsQuery rs, "UPDATE History SET Stamp='" & Now() & "' WHERE ID=" & rs("ID") & ";"
            
        End If

        
        rsClose rs
        Set dbConn = Nothing

        UpdateAutoTypeLists
        
    End If

End Sub

Public Function UpdateAutoTypeLists()
    Dim frm
    Dim ctl

    For Each frm In Forms

        If TypeName(frm) = "frmFTPClientGUI" Then
            SetAutoTypeList frm, frm.setLocation(0)
            SetAutoTypeList frm, frm.setLocation(1)
        Else
            For Each ctl In frm.Controls
                If TypeName(frm) = "SiteInformation" Then
                    SetAutoTypeList frm, ctl.AutoTypeCombo
                
                End If
            Next
        End If

    Next

End Function

Public Sub LoadCache(ByVal sInfo As NTControls22.SiteInformation)
    Dim HostServer As String
    HostServer = sInfo.sHostURL.Text
    If (Right(HostServer, 1) = "/" And HostServer <> "ftp://" And HostServer <> "ftps://") Then HostServer = Left(HostServer, Len(HostServer) - 1)
    
    If (HostServer <> "" And HostServer <> "ftp://" And HostServer <> "ftps://") Then
       
        Dim enc As New NTCipher10.nCode
        
        Dim dbConn As New clsDBConnection
        Dim rs As New ADODB.Recordset

        Dim URL As New NTAdvFTP61.URL
                
        dbConn.rsQuery rs, "SELECT * FROM SiteCache WHERE ParentID=" & dbSettings.CurrentUserID & " AND (HostURL = '" & Replace(HostServer, "'", "''") & "' AND Port = " & sInfo.sPort & " AND SSL = " & sInfo.sSSL.Value & ") OR HostURL = '" & Replace(HostServer, "'", "''") & "';"
        
        If Not rsEnd(rs) Then
    
            If Not (rs("Username") = "") Then sInfo.sUserName.Text = enc.DecryptString(rs("Username"), dbSettings.CryptKey)
            If Not (rs("Password") = "") Then sInfo.sPassword.Text = enc.DecryptString(rs("Password"), dbSettings.CryptKey(sInfo.sUserName.Text))
            sInfo.sPort.Text = CStr(rs("Port"))
            sInfo.sPassive.Value = BoolToCheck(rs("Passive"))
            sInfo.sPortRange.Text = rs("PortRange")
            sInfo.sSSL.Value = rs("SSL")
            sInfo.sAdapter.ListIndex = (rs("Adapter") - 1)
            sInfo.sSavePass.Value = 1
        Else
        
            dbConn.rsQuery rs, "SELECT * FROM SiteCache WHERE ParentID=" & dbSettings.CurrentUserID & " AND (HostURL LIKE '%" & Replace(URL.GetServer(HostServer), "'", "''") & "%' AND Port=" & sInfo.sPort & " AND SSL=" & sInfo.sSSL.Value & ") OR HostURL LIKE '%" & Replace(URL.GetServer(HostServer), "'", "''") & "%';"
           
            If Not rsEnd(rs) Then
        
                If Not (rs("Username") = "") Then sInfo.sUserName.Text = enc.DecryptString(rs("Username"), dbSettings.CryptKey)
                If Not (rs("Password") = "") Then sInfo.sPassword.Text = enc.DecryptString(rs("Password"), dbSettings.CryptKey(sInfo.sUserName.Text))
                sInfo.sPort.Text = CStr(rs("Port"))
                sInfo.sPassive.Value = BoolToCheck(rs("Passive"))
                sInfo.sPortRange.Text = rs("PortRange")
                sInfo.sSSL.Value = rs("SSL")
                sInfo.sAdapter.ListIndex = (rs("Adapter") - 1)
                sInfo.sSavePass.Value = 1
            Else
                
                sInfo.sSavePass.Value = 0
    
            End If

        End If
    
        rsClose rs
        Set dbConn = Nothing
        
        Set enc = Nothing
        
    End If

End Sub
Public Sub SaveCache(ByVal sInfo As NTControls22.SiteInformation)
    Dim HostServer As String
    HostServer = sInfo.sHostURL.Text
    If Right(HostServer, 1) = "/" And HostServer <> "ftp://" And HostServer <> "ftps://" Then HostServer = Left(HostServer, Len(HostServer) - 1)
    
    If (HostServer <> "" And HostServer <> "ftp://" And HostServer <> "ftps://") Then
    
        Dim enc As New NTCipher10.nCode
        Dim dbID As Long
        Dim dbConn As New clsDBConnection
        Dim rs As New ADODB.Recordset
        
        Dim URL As New NTAdvFTP61.URL

        dbConn.rsQuery rs, "DELETE FROM SiteCache WHERE ParentID=" & dbSettings.CurrentUserID & " AND (HostURL='" & Replace(HostServer, "'", "''") & "' AND Port=" & sInfo.sPort & " AND SSL=" & sInfo.sSSL.Value & ") OR HostURL='" & Replace(HostServer, "'", "''") & "';"
    
        If sInfo.sSavePass.Value = 1 Then
            If sInfo.sUserName.Text <> "" And sInfo.sPassword.Text <> "" Then
            dbConn.rsQuery rs, "INSERT INTO SiteCache (ParentID, HostURL, Username, Password, Port, Passive, PortRange, Adapter, DateAltered, SSL) " & _
                    "VALUES (" & dbSettings.CurrentUserID & ", '" & Replace(HostServer, "'", "''") & "', '" & enc.EncryptString(Replace(sInfo.sUserName.Text, "'", "''"), dbSettings.CryptKey) & "', '" & enc.EncryptString(Replace(sInfo.sPassword.Text, "'", "''"), dbSettings.CryptKey(Replace(sInfo.sUserName.Text, "'", "''"))) & "', " & sInfo.sPort.Text & ", " & CheckToBool(sInfo.sPassive.Value) & ", '" & Replace(sInfo.sPortRange.Text, "'", "''") & "', " & (sInfo.sAdapter.ListIndex + 1) & ", '" & Now & "', " & sInfo.sSSL.Value & ");"
            End If
        Else
        
            dbConn.rsQuery rs, "DELETE FROM SiteCache WHERE ParentID=" & dbSettings.CurrentUserID & " AND (HostURL LIKE '%" & Replace(URL.GetServer(HostServer), "'", "''") & "%' AND Port=" & sInfo.sPort & " AND SSL=" & sInfo.sSSL.Value & ") OR HostURL LIKE '%" & Replace(URL.GetServer(HostServer), "'", "''") & "%';"
            
        End If
    
        rsClose rs
        Set dbConn = Nothing
        
        Set enc = Nothing
        Set URL = Nothing
        
    End If

End Sub
