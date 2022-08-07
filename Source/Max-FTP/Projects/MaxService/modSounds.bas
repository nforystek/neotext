Attribute VB_Name = "modSounds"
Option Explicit
'TOP DOWN
Option Private Module
Private TransferCount As Long
Public SoundNotify As Boolean
Public clsSounds As clsSoundSet

Public Sub SoundItteration()
    If SoundNotify Then
        If (clsSounds Is Nothing) Then Set clsSounds = New clsSoundSet

        TransferCount = 0
        
        If (AllSchedules.Count > 0) Then
            Dim sc As clsSchedule
            Dim op As clsOperation
            For Each sc In AllSchedules
                For Each op In sc.Operations
                    If Not (TransferCount = -1) Then
                        If op.ErrorOccur Then
                            TransferCount = -1
                        ElseIf op.Transfering Then
                            TransferCount = TransferCount + 1
                        End If
                    End If
                Next
            Next
        End If
        
        If TransferCount = -1 Then
            PlayContinuousErrorNote
        ElseIf TransferCount > 0 Then
            PlayOneTickPerTransfer
        Else
            clsSounds.SoundHalt False
            clsSounds.SoundTick 0
        End If
    
    Else
        clsSounds.SoundHalt False
        clsSounds.SoundTick 0
    End If
End Sub

Public Sub PlayContinuousErrorNote()
    clsSounds.SoundHalt True
End Sub

Public Sub PlayOneTickPerTransfer()
    clsSounds.SoundHalt False
    clsSounds.SoundTick TransferCount
End Sub

Public Sub PlayStartTransferSound()
    If SoundNotify Then
        If (clsSounds Is Nothing) Then Set clsSounds = New clsSoundSet
        clsSounds.SoundPlay
    End If
End Sub

Public Sub PlayStopTransferSound()
    If SoundNotify Then
        If (clsSounds Is Nothing) Then Set clsSounds = New clsSoundSet
        If TransferCount = 1 Then clsSounds.SoundTick 0
        clsSounds.SoundStop
    End If
End Sub
