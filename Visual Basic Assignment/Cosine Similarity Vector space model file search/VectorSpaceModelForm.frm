VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VectorSpaceModelSearchEngine 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   4560
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Display/Hide Vector Model  "
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   16
      Top             =   5280
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   14640
      TabIndex        =   12
      Text            =   "Files Retrieval Order:"
      Top             =   240
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   11280
      TabIndex        =   7
      Top             =   2760
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   7920
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   5
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid table 
      Height          =   4575
      Left            =   360
      TabIndex        =   4
      Top             =   6000
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   12
      BackColor       =   12640511
      BackColorFixed  =   12648384
      BackColorSel    =   16776960
      BackColorBkg    =   16744576
      GridColor       =   4210752
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Franklin Gothic Demi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3000
      TabIndex        =   0
      Top             =   2760
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "LOWEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   18000
      TabIndex        =   15
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "HIGHEST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   18000
      TabIndex        =   14
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   " Match to Match"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   18120
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Caption         =   "Directory:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "File List:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "FILE SEARCH MECHANISM                  BASED ON                 VECTOR SPACE MODELLING"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "QUERRY (keyword):"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   2460
      Width           =   2295
   End
End
Attribute VB_Name = "VectorSpaceModelSearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer   'local
Dim kp As Integer   'display and hide
Dim qry As String    'querry string
Dim finaltext As String   'collection of all document text
Dim donelen As Integer  'Non repeated words length
Dim arrlen As Integer   'collection of split words in final text
Dim qrylen As Integer   'split words length in querry
Dim terms(1000, 1000) As Double    'table of vector model
Dim arr(10000) As String      'split words of final text
Dim querry(100) As String     'split words of querry
Dim done(10000) As String     'split words of non repeated array
Dim fle(10000, 10000) As String       'file related
Dim flen(10000) As Integer           'file related
Dim filecontent(10000) As String     'file related
Dim rowlen As Integer                'Total columns in matrix

Private Sub Command1_Click()
    Dim i As Integer
    Dim j As Integer
    
    '1. Read File Content and Querry
    
    Call ReadDataAndQuerry
    
    '2. get all the collection of document strings in final text
    
    Call ConcatFile
    
    '3. initialisations
    
    donelen = arrlen = qrylen = 0
    
    '4. split(finaltext," ") and get length of splitted string
    
    Call wordsintext(finaltext, arr, arrlen)
    
    '5. split(querry," ") and get length of splitted string
    
    Call wordsintext(qry, querry, qrylen)
    
    '6. sort in alphabetic order A - Z
    
    Call sort(arr, arrlen)
    Call sort(querry, qrylen)
    
    '7. Remove redundancy of duplicate words in split array
    
    Call remduplication(done, donelen, arr, arrlen)
    
    '8. quantify terms frequency
    
    Call termfrequency
    
    '9. calculate the document frequency
    
    Call CalcDFI
    
    '10. calculate the document frequency ratio to number of documents
    
    Call CalcDDFI
    
    '11. calculate the IDF Value
    
    Call CalcIDF
    
    '12. calculate the weights of each document Corresponding to the querry
    
    Call Calcweights
            
    '13. Table Orientation of Vector Space Model Matrix
    
    Call TableOrientation
    
    '14. Calculate euclidean length of vector
    
    Call euclideanvector
    
    '15. Calculate Dot vector Values
    
    Call DotProduct
    
    '16. Calculate Cosine Values To determine highest match
    
    Call CalcCOS
    
    '17. Print Vector Model Data into Table Matrix
    
    Call PrintData
    
    '18. File Retrieval Order
    
    Call FileRetrieve
    
End Sub

Private Sub Command2_Click()
    Text6.text = ""
End Sub

Private Sub Command3_Click()
    'MsgBox (kp)
    If kp = 0 Then
    table.Visible = True
    kp = 1
    Else
    table.Visible = False
    kp = 0
    End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Public Sub wordsintext(text As String, ByRef arr() As String, ByRef n As Integer)
    Dim ln As Integer
    Dim i As Integer
    Dim j As Integer
    i = 1
    j = 0
    text = Trim(text)
    ln = Len(text)
    Dim cnt As Integer
    cnt = 0
    Do While i <= ln
        Dim ch As String
        ch = Mid(text, i, 1)
        If ch <> " " And ch <> "." Then
            cnt = cnt + 1
        Else
            If Mid(text, i - cnt, cnt) <> "" And Mid(text, i - cnt, cnt) <> " " And Mid(text, i - cnt, cnt) <> "  " Then
            arr(j) = Mid(text, i - cnt, cnt)
            j = j + 1
            cnt = 0
            End If
        End If
        i = i + 1
    Loop
    If Mid(text, i - cnt, cnt) <> "" And Mid(text, i - cnt, cnt) <> " " And Mid(text, i - cnt, cnt) <> "  " Then
    arr(j) = Mid(text, i - cnt, cnt)
    End If
    n = j
End Sub


Public Sub wordsintext1(text As String, ByVal k As Integer, ByRef n As Integer)
    Dim ln As Integer
    Dim i As Integer
    Dim j As Integer
    i = 1
    j = 0
    text = Trim(text)
    ln = Len(text)
    Dim cnt As Integer
    cnt = 0
    Do While i <= ln
        Dim ch As String
        ch = Mid(text, i, 1)
        If ch <> " " And ch <> "." Then
            cnt = cnt + 1
        Else
            If Mid(text, i - cnt, cnt) <> "" And Mid(text, i - cnt, cnt) <> " " And Mid(text, i - cnt, cnt) <> "  " Then
            fle(k, j) = Mid(text, i - cnt, cnt)
            j = j + 1
            cnt = 0
            End If
        End If
        i = i + 1
    Loop
    If Mid(text, i - cnt, cnt) <> "" And Mid(text, i - cnt, cnt) <> " " And Mid(text, i - cnt, cnt) <> "  " Then
    fle(k, j) = Mid(text, i - cnt, cnt)
    End If
    n = j
End Sub

Public Sub sort(ByRef arr() As String, n As Integer)
Dim i As Integer
    Dim j As Integer
    For i = 0 To n - 1
        For j = 0 To n - i - 1
        If arr(j) <> "" And arr(j) <> " " And arr(j) <> "  " And arr(j + 1) <> "" And arr(j + 1) <> " " And arr(j + 1) <> "  " Then
            If Len(arr(j)) > Len(arr(j + 1)) Then
                Dim temp As String
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        End If
        Next j
    Next i
    For i = 0 To n - 1
        For j = 0 To n - i - 1
        If arr(j) <> "" And arr(j) <> " " And arr(j) <> "  " And arr(j + 1) <> "" And arr(j + 1) <> " " And arr(j + 1) <> "  " Then
            If Asc(arr(j)) > Asc(arr(j + 1)) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        End If
        Next j
    Next i
End Sub

Public Sub remduplication(ByRef done() As String, ByRef j As Integer, arr() As String, n As Integer)
    Dim i As Integer
    For i = 0 To n
        If i = 0 Then
            j = 0
            done(j) = arr(i)
            j = j + 1
        End If                              ' a a b c c d e f f
            Dim k As Integer
            For k = 0 To j
                If arr(i) = done(k) Then
                    Exit For
                End If
            Next k
            If k = j + 1 Then
                done(j) = arr(i)
                j = j + 1
            End If
    Next i
End Sub

Public Sub termfrequency()
Dim i As Integer
    Dim j As Integer
    For i = 0 To donelen - 1
        For j = 0 To File1.ListCount - 1
        Dim k As Integer
            For k = 0 To flen(j)
                If done(i) = fle(j, k) Then
                    terms(i, j) = terms(i, j) + 1
                End If
            Next k
        Next j
        If j = File1.ListCount Then
                For k = 0 To qrylen
                    If done(i) = querry(k) Then
                        terms(i, j) = terms(i, j) + 1
                    End If
                Next k
            End If
    Next i
End Sub


Public Sub sortnum(ByRef ar() As Double, n As Integer, ByRef cos() As String)
Dim temp As String
Dim tempk As Double
    For i = 0 To n - 2
        For j = 0 To n - i - 2
            If ar(j) < ar(j + 1) Then
                temp = cos(j)
                cos(j) = cos(j + 1)
                cos(j + 1) = temp
                tempk = ar(j)
                ar(j) = ar(j + 1)
                ar(j + 1) = tempk
            End If
        Next j
    Next i
End Sub

Private Sub Form_Load()
    File1.Path = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\docsam"
    Dir1.Path = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\docsam"
    table.Visible = False
    kp = 0
End Sub

Public Sub CalcDFI()
Dim i As Integer
Dim j As Integer
    For i = 0 To donelen - 1 'DFI
        terms(i, File1.ListCount + 1) = 0
        For j = 0 To File1.ListCount - 1
            If terms(i, j) <> 0 Then
            terms(i, File1.ListCount + 1) = terms(i, File1.ListCount + 1) + 1
            End If
        Next j
    Next i
End Sub

Public Sub CalcDDFI()
Dim i As Integer
    For i = 0 To donelen - 1 'D/DFI D:no of documents 3
        terms(i, File1.ListCount + 2) = Round(3 / terms(i, File1.ListCount + 1), 2)
    Next i
End Sub

Public Sub CalcIDF()
Dim i As Integer
    For i = 0 To donelen - 1 'IDF = Log(D/dfi)
        terms(i, File1.ListCount + 3) = Round(Log(terms(i, File1.ListCount + 2)) / Log(10), 4)
    Next i
End Sub

Public Sub Calcweights()
Dim i As Integer
Dim j As Integer
    rowlen = ((File1.ListCount + 1) * 2) + 3
    For i = 0 To donelen - 1
        For j = File1.ListCount + 4 To rowlen - 1
            terms(i, j) = Round(terms(i, j - 4 - File1.ListCount) * terms(i, File1.ListCount + 3), 4)
        Next j
    Next i
End Sub

Public Sub TableOrientation()
Dim i As Integer
Dim j As Integer
    table.Rows = donelen + 4
    table.Cols = rowlen + 1
    table.TextMatrix(0, 0) = " "
    For j = 1 To File1.ListCount
        table.TextMatrix(0, j) = "D" + Str(j)
    Next j
    table.TextMatrix(0, File1.ListCount + 1) = "Q"
    table.TextMatrix(0, File1.ListCount + 2) = "DFI"
    table.TextMatrix(0, File1.ListCount + 3) = "D/DFI"
    table.TextMatrix(0, File1.ListCount + 4) = "IDF"
    i = 1
    For j = File1.ListCount + 5 To rowlen - 1
    table.TextMatrix(0, j) = "WD" + Str(i)
    i = i + 1
    Next j
    table.TextMatrix(0, rowlen) = "WDQ"
    For i = 1 To donelen
    table.TextMatrix(i, 0) = done(i - 1)
    Next i
    table.TextMatrix(donelen + 1, 0) = "eucledian vector"
    table.TextMatrix(donelen + 2, 0) = "DOT product"
    table.TextMatrix(donelen + 3, 0) = "cos theta"
End Sub

Public Sub euclideanvector()
Dim i As Integer
Dim j As Integer
    For j = File1.ListCount + 4 To rowlen - 1
    terms(donelen, j) = 0
        For i = 0 To donelen - 1
            terms(donelen, j) = Round(terms(donelen, j) + (terms(i, j) ^ 2), 4)
        Next i
    terms(donelen, j) = Round(Sqr(terms(donelen, j)), 4)
    Next j
End Sub

Public Sub DotProduct()
Dim i As Integer
Dim j As Integer
    For j = File1.ListCount + 4 To rowlen - 2
        terms(donelen + 1, j) = 0
        For i = 0 To donelen - 1
            terms(donelen + 1, j) = terms(donelen + 1, j) + terms(i, j) * terms(i, rowlen - 1)
        Next i
    terms(donelen + 1, j) = Round(terms(donelen + 1, j), 4)
    Next j
End Sub

Public Sub CalcCOS()
Dim j As Integer
    For j = File1.ListCount + 4 To rowlen - 2
        terms(donelen + 2, j) = Round(terms(donelen + 1, j) / (terms(donelen, rowlen - 1) * terms(donelen, j)), 4)
    Next j
End Sub

Public Sub PrintData()
Dim i As Integer
Dim j As Integer
    For i = 1 To donelen + 3
        For j = 1 To rowlen
            table.TextMatrix(i, j) = Str(terms(i - 1, j - 1))
        Next j
    Next i
End Sub

Public Sub FileRetrieve()
    Dim i As Integer
    Dim j As Integer
    Dim cosval(10000) As Double
    Dim cosstring(10000) As String
    i = 0
    For j = File1.ListCount + 4 To rowlen - 2
        cosval(i) = terms(donelen + 2, j)
        cosstring(i) = table.TextMatrix(0, j + 1)
        i = i + 1
    Next j
    Call sortnum(cosval, File1.ListCount, cosstring)
    For i = 0 To File1.ListCount - 1
        Dim y As Integer
        y = Val(Mid(cosstring(i), 4, 1))
        Combo1.AddItem File1.List(y - 1), i
    Next i
End Sub


Public Sub ConcatFile()
Dim i As Integer
    For i = 0 To File1.ListCount - 1
    finaltext = finaltext + filecontent(i) + " "
    Next i
End Sub

Public Sub ReadDataAndQuerry()
    Dim i As Integer
    For i = 0 To File1.ListCount - 1
        Dim strFile As String
        'With CD
        '    .ShowOpen
        '   Text1.Text = .FileName
        'End With
        Open File1.Path + "\" + File1.List(i) For Input As #1
        Do While Not (EOF(1))
            Line Input #1, strFile
            filecontent(i) = filecontent(i) + " " + strFile
        Loop
        Call wordsintext1(filecontent(i), i, flen(i))
        Close #1
    Next i
    qry = Text6.text
End Sub


