Attribute VB_Name = "Module1"
Option Explicit
Public Instances As Note


Private Lists As New Notes

Private Note As Note

Public Sub Main()

     Set Note = Lists.CreateList
    
    
    
    Debug.Print "No List: " & Note.Handle
    Note.Insert
    Debug.Print "Insert: " & Note.Handle
    
    Note.Insert
    Debug.Print "Insert: " & Note.Handle
    Note.Insert
    Debug.Print "Insert: " & Note.Handle
    
    
    Set Note = Note.Forth
    Debug.Print "Prior: " & Note.Handle
    
    
    
    Lists.DisposeList Note
    Set Note = Nothing

End Sub
