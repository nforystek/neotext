VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN
Option Compare Binary

Public Enum MemoryScopes
    Heap = 0
    Locking = 1
    [Local] = 2
    [Global] = 4
End Enum

Public Function Address(Optional ByVal Offset As Long = 0) As Long ' _
Retrieves the address numberical value where the memory this object is allocated at.
Attribute Address.VB_Description = "Retrieves the address numberical value where the memory this object is allocated at."

End Function

Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, but does not change the scope of the object cloning the Origin, to do so, use Reset before cloning.
Attribute Clone.VB_Description = "Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, but does not change the scope of the object cloning the Origin, to do so, use Reset before cloning."

End Sub

Public Sub Post(ByRef RHS() As Byte) ' _
Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Concat()
Attribute Post.VB_Description = "Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Concat()"

End Sub

Public Sub Concat(ByRef RHS() As Byte) ' _
Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Post()
Attribute Concat.VB_Description = "Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Post()"

End Sub

Public Property Get Length() As Long ' _
Gets the amount of bytes with in this instance of the stream object.
Attribute Length.VB_Description = "Gets the amount of bytes with in this instance of the stream object."

End Property

Public Property Let Length(ByVal RHS As Long) ' _
Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes.
Attribute Length.VB_Description = "Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes."

End Property

Public Sub PlaceAt(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
Places contents of memory byte array RHS into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted.
Attribute PlaceAt.VB_Description = "Places contents of memory byte array RHS into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted."

End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte() ' _
Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve.
Attribute Partial.VB_Description = "Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve."

End Function

Public Sub Prepend(ByRef RHS() As Byte) ' _
Places RHS bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in RHS.
Attribute Prepend.VB_Description = "Places RHS bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in RHS."

End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left.
Attribute Pull.VB_Description = "Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left."

End Sub

Public Sub Push(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
Attribute Push.VB_Description = "Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right."

End Sub

Public Sub Reset(Optional ByRef Scopes As MemoryScopes) ' _
Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use.
Attribute Reset.VB_Description = "Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use."

End Sub

Public Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Long ' _
Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width.
Attribute Pass.VB_Description = "Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width."

End Function

Public Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Long ' _
Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width.
Attribute Poll.VB_Description = "Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width."

End Function

Public Sub Pinch(ByVal Offset As Long, ByVal Width As Long) ' _
Removes a section of the stream that is in the middle having portions of the stream before and after it resulting as continuous, least the pinched Width.
Attribute Pinch.VB_Description = "Removes a section of the stream that is in the middle having portions of the stream before and after it resulting as continuous, least the pinched Width."

End Sub

Public Sub Pyramid(ByRef SubStream As IStream, ByVal Offset As Long, ByVal Width As Long) ' _
Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream.
Attribute Pyramid.VB_Description = "Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream."

End Sub

Public Function Peek(Optional ByVal Offset As Long = 0) As Byte ' _
Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset.
Attribute Peek.VB_Description = "Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset."

End Function



