VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Jpk Archiver"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Extract"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "List files"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "c:\windows\desktop\"
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "c:\windows\desktop\"
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Archive:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File to add:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemLength As Long
Dim ItemString As String
Dim ItemNumber(0 To 1) As Integer

Dim BytesExtract As String
Dim BytesAdd As String

Dim ItemBinary As String
Dim Position As Long
Dim LastPosition As Long

Dim FileListStart As Long
Dim FilePosition As Long
Dim ExitDo As Boolean

Dim PutLength As String
Dim PutPosition As Long

Function JpkAdd(JpkFile As String, FileName As String, AddName As String) As Boolean

    On Error GoTo FinaliseError

    AddName = AddName & Chr(0)
    
    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open FileName For Binary As #ItemNumber(1)
            PutLength = LOF(ItemNumber(1)) & Chr(0)
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, AddName
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, PutLength
            PutPosition = LOF(ItemNumber(0))
            If LOF(ItemNumber(1)) > 1000000 Then
                Position = -999999
                Do
                    Position = Position + 1000000
                    If Position + 999999 > LOF(ItemNumber(1)) Then BytesAdd = String(LOF(ItemNumber(1)) - Position + 1, Chr$(0)) Else BytesAdd = String(1000000, Chr$(0))
                    Get ItemNumber(1), Position, BytesAdd
                    Put ItemNumber(0), PutPosition + 1, BytesAdd
                    PutPosition = LOF(ItemNumber(0))
                Loop Until Position + 999999 > LOF(ItemNumber(1))
            Else
                BytesAdd = String(LOF(ItemNumber(1)), Chr$(0))
                Get ItemNumber(1), , BytesAdd
                Put ItemNumber(0), PutPosition + 1, BytesAdd
            End If
        Close ItemNumber(1)
    Close #ItemNumber(0)
    JpkAdd = True
    Exit Function
    
FinaliseError:
    JpkAdd = False

End Function

Function JpkList(JpkFile As String, ListItem As ListBox) As Boolean

    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        Position = 1
        Do
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            Position = Position + Len(ItemString) + 1
            ListItem.AddItem ItemString
            
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            ItemLength = CLng(ItemString)
            Position = Position + Len(ItemString) + ItemLength + 1
        Loop Until Position > LOF(ItemNumber(0))
    Close #ItemNumber(0)
    JpkList = True
    Exit Function
    
FinaliseError:
    JpkList = False

End Function

Function JpkExtract(JpkFile As String, FileName As String, Destination As String) As Boolean

    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open Destination For Binary As ItemNumber(1)
            Position = 1
            ExitDo = False
            Do
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                Position = Position + Len(ItemString) + 1
                If LCase(ItemString) = LCase(FileName) Then ExitDo = True
                
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                ItemLength = CLng(ItemString)
                Position = Position + Len(ItemString) + ItemLength + 1
                If ExitDo = True Then Exit Do
            Loop Until Position > LOF(ItemNumber(0))
            
            FileListStart = Position - ItemLength
            If ItemLength > 1000000 Then
                FilePosition = -999999
                Do
                    FilePosition = FilePosition + 1000000
                    If FilePosition + 999999 > ItemLength Then BytesExtract = Space(ItemLength - FilePosition + 1) Else BytesExtract = Space(1000000)
                    Get ItemNumber(0), FileListStart, BytesExtract
                    Put ItemNumber(1), FilePosition, BytesExtract
                    FileListStart = FileListStart + Len(BytesExtract)
                Loop Until FilePosition + 999999 > LOF(ItemNumber(1))
            Else
                BytesExtract = Space(ItemLength)
                Get ItemNumber(0), Position - ItemLength, BytesExtract
                Put ItemNumber(1), 1, BytesExtract
            End If
        Close ItemNumber(1)
    Close ItemNumber(0)
    JpkExtract = True
    Exit Function
    
FinaliseError:
    JpkExtract = False

End Function

Private Sub Command1_Click()
    JpkAdd Text2, Text1, GetFileName(Text1)
End Sub

Private Sub Command2_Click()
    JpkList Text2, List1
End Sub

Private Sub Command3_Click()
    JpkExtract Text2, List1.Text, Text4
End Sub

Private Sub List1_Click()
    Text4 = GetPath(Text2) & List1.Text
End Sub
