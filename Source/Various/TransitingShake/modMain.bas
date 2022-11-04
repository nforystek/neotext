Attribute VB_Name = "modMain"
#Const True = -1
#Const False = 0
#Const modMain = -1
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Sub RandomPause()
    Randomize
    Sleep (Rnd * 1000) + (Rnd * 1000) + (Rnd * 1000)
End Sub
Public Sub Main()
    Dim Phrase As String
    
    Dim client As Handshake
    Dim server As Handshake

    Dim client2 As Handshake
    Dim server2 As Handshake
    
    
    Debug.Print "In test 1 through 4 are variant displays of success and failures to checkum."
    Debug.Print "These are depnding on the inital checksums for each client and server callee"
    Debug.Print "arguments, to enforce following checksum calls that must match in patterns."
    Debug.Print
    Debug.Print "Test 5 shows how the server and client can do the same thing effectively with"
    Debug.Print "out actually using the other objects properties, if for instance communication"
    Debug.Print "is being transmitted, the instance of that transmit may act as timing checksum."
    Debug.Print
    Debug.Print "Because failure may change a success rate of future calls, there is a Rollback."
    Debug.Print

    Debug.Print "Beginning Test 1"

        
    Set client = New Handshake
    Set server = New Handshake
    RandomPause
    

    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, client.Seed), "Success", "Failure")
    
    Debug.Print "Server Checksum " & IIf(server.Checksum(client.Para, server.Seed), "Success", "Failure")

    RandomPause
    
    Debug.Print "Client Checksum " & IIf(client.Checksum(client.Para, server.Seed), "Success", "Failure")

    RandomPause
    Debug.Print "Client Checksum " & IIf(client.Checksum(client.Para, client.Seed), "Success", "Failure")

    RandomPause
    Debug.Print "Server Checksum " & IIf(server.Checksum(server.Para, client.Seed), "Success", "Failure")

    RandomPause
    Debug.Print "Client Checksum " & IIf(server.Checksum(server.Para, server.Seed), "Success", "Failure")


    Set client = Nothing
    Set server = Nothing
    

    Debug.Print "Beginning Test 2"
    
    Set client = New Handshake 'creating the objects start the seeding internalls until
    
    Set server = New Handshake 'brought out through active use or remaining passive use
    
    
    RandomPause
    
    
    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, client.Seed), "Success", "Failure")
    
    RandomPause
    
    Debug.Print "Server Checksum " & IIf(server.Checksum(client.Para, client.Seed), "Success", "Failure")
           
    
    RandomPause
    Debug.Print "Server Checksum " & IIf(server.Checksum(client.Para, server.Seed), "Success", "Failure")
    
    RandomPause
    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, client.Seed), "Success", "Failure")
    
    
    Set client = Nothing
    Set server = Nothing
    

    Debug.Print "Beginning Test 3"
    
    Set client = New Handshake 'creating the objects start the seeding internalls until
    
    Set server = New Handshake 'brought out through active use or remaining passive use
    
    
    RandomPause
    
    
    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, server.Seed), "Success", "Failure")
    
    RandomPause
    
    Debug.Print "Server Checksum " & IIf(server.Checksum(server.Para, server.Seed), "Success", "Failure")
    RandomPause
    Debug.Print "Client Checksum " & IIf(client.Checksum(client.Para, client.Seed), "Success", "Failure")
    
    RandomPause
    Debug.Print "Server Checksum " & IIf(server.Checksum(client.Para, server.Seed), "Success", "Failure")
    

    RandomPause
    
    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, client.Seed), "Success", "Failure")
           
    
    Set client = Nothing
    Set server = Nothing
    
    Debug.Print "Beginning Test 4"
    
    Set client = New Handshake 'creating the objects start the seeding internalls until
    
    Set server = New Handshake 'brought out through active use or remaining passive use
    
    
    RandomPause
    
    
    Debug.Print "Client Initials " & IIf(client.Checksum(server.Para, client.Seed), "Success", "Failure")
    
    RandomPause
    
    Debug.Print "Server Initials " & IIf(server.Checksum(client.Para, client.Seed), "Success", "Failure")
           
    
    RandomPause
    Debug.Print "Server Checksum " & IIf(server.Checksum(server.Para, client.Seed), "Success", "Failure")
    
    RandomPause
    Debug.Print "Client Checksum " & IIf(client.Checksum(server.Para, server.Seed), "Success", "Failure")
    
    
    Set client = Nothing
    Set server = Nothing
    
    
    Debug.Print "Beginning Test 5"
        
        
    Set client = New Handshake
        
    Set server = New Handshake

    Set client2 = New Handshake
        
    Set server2 = New Handshake
    
    RandomPause
    

    Debug.Print "Client Initials " & IIf(client.Checksum(client2.Para, client.Seed), "Success", "Failure")
    Debug.Print "Server Initials " & IIf(server.Checksum(server2.Para, server.Seed), "Success", "Failure")
    

    Debug.Print "Client2 Checksum " & IIf(client2.Checksum(client2.Para, client.Seed), "Success", "Failure")
    Debug.Print "Server2 Checksum " & IIf(server2.Checksum(server2.Para, server.Seed), "Success", "Failure")

    RandomPause
    Debug.Print "Client Checksum " & IIf(client.Checksum(client2.Para, client.Seed), "Success", "Failure")
    Debug.Print "Server Checksum " & IIf(server2.Checksum(server.Para, server2.Seed), "Success", "Failure")

    Set client = Nothing
    Set server = Nothing
    Set client2 = Nothing
    Set server2 = Nothing
End Sub

