VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid table 
      Height          =   855
      Left            =   480
      TabIndex        =   10
      Top             =   4800
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   615
      Left            =   9840
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "search"
      Height          =   615
      Left            =   7800
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   2960
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   0
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Document content 1:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   2220
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Document content 2:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   480
      TabIndex        =   6
      Top             =   3015
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Document content 3:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   3780
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "querry:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1380
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i, j As Integer
    Dim t1 As String
    Dim t2 As String
    Dim t3 As String
    Dim qry As String
    Dim finaltext As String
    Dim donelen As Integer
    Dim arrlen As Integer
    Dim qrylen As Integer
    Dim terms(10000, 100) As Double
    Dim arr(100000) As String
    Dim querry(100) As String
    Dim done(1000000) As String
    Dim d1len As Integer
    Dim d2len As Integer
    Dim d3len As Integer
    Dim d1(100) As String
    Dim d2(100) As String
    Dim d3(100) As String
    t1 = Text1.text
    t2 = Text4.text
    t3 = Text5.text
    qry = Text6.text
    finaltext = t1 + " " + t2 + " " + t3
    donelen = arrlen = qrylen = 0
    'text = InputBox("enter the text:")
    'MsgBox ("the position is " + Str(InStr(1, text, " ")))
    Call wordsintext(finaltext, arr, arrlen)
    Call wordsintext(qry, querry, qrylen)
    For i = 0 To qrylen
    MsgBox (querry(i))
    Next i
    Call wordsintext(t1, d1, d1len)
    Call wordsintext(t2, d2, d2len)
    Call wordsintext(t3, d3, d3len)
    Call sort(arr, arrlen)
    Call sort(querry, qrylen)
    Call remduplication(done, donelen, arr, arrlen)
    'MsgBox (n) here it prints n 1 value less as array starts from 0
    For i = 0 To donelen - 1
        For j = 0 To 3
            terms(i, j) = 0
        Next j
    Next i
    For i = 0 To donelen - 1
        For j = 0 To 3
        Dim k As Integer
            If j = 0 Then
                For k = 0 To d1len
                    If done(i) = d1(k) Then
                        terms(i, j) = terms(i, j) + 1
                    End If
                Next k
            End If
            If j = 1 Then
                For k = 0 To d2len
                    If done(i) = d2(k) Then
                        terms(i, j) = terms(i, j) + 1
                    End If
                Next k
            End If
            If j = 2 Then
                For k = 0 To d3len
                    If done(i) = d3(k) Then
                        terms(i, j) = terms(i, j) + 1
                    End If
                Next k
            End If
            If j = 3 Then
                For k = 0 To qrylen
                    If done(i) = querry(k) Then
                        terms(i, j) = terms(i, j) + 1
                    End If
                Next k
            End If
        Next j
    Next i
    For i = 0 To donelen - 1 'DFI
        terms(i, 4) = 0
        For j = 0 To 2
            If terms(i, j) <> 0 Then
                terms(i, 4) = terms(i, 4) + 1
            End If
        Next j
    Next i
    For i = 0 To donelen - 1 'D/DFI D:no of documents 3
        terms(i, 5) = Round(3 / terms(i, 4), 2)
    Next i
    For i = 0 To donelen - 1 'IDF = Log(D/dfi)
        terms(i, 6) = Round(Log(terms(i, 5)) / Log(10), 4)
    Next i
    For i = 0 To donelen - 1 'weights of d1
        terms(i, 7) = Round(terms(i, 0) * terms(i, 6), 4)
    Next i
        For i = 0 To donelen - 1 'weights of d2
        terms(i, 8) = Round(terms(i, 1) * terms(i, 6), 4)
    Next i
        For i = 0 To donelen - 1 'weights of d3
        terms(i, 9) = Round(terms(i, 2) * terms(i, 6), 4)
    Next i
        For i = 0 To donelen - 1 'weights of Q
        terms(i, 10) = Round(terms(i, 3) * terms(i, 6), 4)
    Next i
    'For i = 0 To donelen - 1
    '    For j = 0 To 10
    '        Print (Str(terms(i, j)) + "    ");
    '    Next j
    '    Print (vbNewLine)
    'Next i
    table.Rows = donelen + 4
    table.Height = (donelen + 3) * 350
    table.TextMatrix(0, 0) = " "
    table.TextMatrix(0, 1) = "D1"
    table.TextMatrix(0, 2) = "D2"
    table.TextMatrix(0, 3) = "D3"
    table.TextMatrix(0, 4) = "Q"
    table.TextMatrix(0, 5) = "DFI"
    table.TextMatrix(0, 6) = "D/DFI"
    table.TextMatrix(0, 7) = "IDF"
    table.TextMatrix(0, 8) = "WD1"
    table.TextMatrix(0, 9) = "WD2"
    table.TextMatrix(0, 10) = "WD3"
    table.TextMatrix(0, 11) = "WDQ"
    For i = 1 To donelen
    table.TextMatrix(i, 0) = done(i - 1)
    Next i
    table.TextMatrix(donelen + 1, 0) = "eucledian vector"
    table.TextMatrix(donelen + 2, 0) = "DOT product"
    table.TextMatrix(donelen + 3, 0) = "cos theta"
    'round(,)                                                                             'euclidean length of vector
    terms(donelen, 7) = 0  'assign |d1| = 0
    For i = 0 To donelen - 1
        terms(donelen, 7) = Round(terms(donelen, 7) + (terms(i, 7) ^ 2), 4)
    Next i
    terms(donelen, 7) = Round(Sqr(terms(donelen, 7)), 4)
    
    terms(donelen, 8) = 0  'assign |d2| = 0
    For i = 0 To donelen - 1
        terms(donelen, 8) = Round(terms(donelen, 8) + (terms(i, 8) ^ 2), 4)
    Next i
    terms(donelen, 8) = Round(Sqr(terms(donelen, 8)), 4)

    terms(donelen, 9) = 0  'assign |d3| = 0
    For i = 0 To donelen - 1
        terms(donelen, 9) = Round(terms(donelen, 9) + (terms(i, 9) ^ 2), 4)
    Next i
    terms(donelen, 9) = Round(Sqr(terms(donelen, 9)), 4)

    terms(donelen, 10) = 0  'assign |Q| = 0
    For i = 0 To donelen - 1
        terms(donelen, 10) = Round(terms(donelen, 10) + (terms(i, 10) ^ 2), 4)
    Next i
    terms(donelen, 10) = Round(Sqr(terms(donelen, 10)), 4)
    
    terms(donelen + 1, 7) = 0                                                         'dot product
    For i = 0 To donelen - 1
            terms(donelen + 1, 7) = terms(donelen + 1, 7) + terms(i, 7) * terms(i, 10)
    Next i
    terms(donelen + 1, 7) = Round(terms(donelen + 1, 7), 4)
    
    For i = 0 To donelen - 1
            terms(donelen + 1, 8) = terms(donelen + 1, 8) + terms(i, 8) * terms(i, 10)
    Next i
    terms(donelen + 1, 8) = Round(terms(donelen + 1, 8), 4)
    
    For i = 0 To donelen - 1
            terms(donelen + 1, 9) = terms(donelen + 1, 9) + terms(i, 9) * terms(i, 10)
    Next i
    terms(donelen + 1, 9) = Round(terms(donelen + 1, 9), 4)
    
    terms(donelen + 2, 7) = Round(terms(donelen + 1, 7) / (terms(donelen, 10) * terms(donelen, 7)), 4) 'cosine theta of |d1|
    terms(donelen + 2, 8) = Round(terms(donelen + 1, 8) / (terms(donelen, 10) * terms(donelen, 8)), 4) 'cosine theta of |d1|
    terms(donelen + 2, 9) = Round(terms(donelen + 1, 9) / (terms(donelen, 10) * terms(donelen, 9)), 4) 'cosine theta of |d1|
    
    For i = 1 To donelen + 3
        For j = 1 To 11
            table.TextMatrix(i, j) = Str(terms(i - 1, j - 1))
        Next j
    Next i
    MsgBox (donelen)
    'MsgBox (Format(2, "0.0000"))
End Sub

Public Sub wordsintext(text As String, ByRef arr() As String, ByRef n As Integer)
    Dim ln As Integer
    i = 1
    j = 0
    ln = Len(text)
    Dim cnt As Integer
    cnt = 0
    Do While i <= ln
        Dim ch As String
        ch = Mid(text, i, 1)
        If ch <> " " Then
            cnt = cnt + 1
        Else
            arr(j) = Mid(text, i - cnt, cnt)
            j = j + 1
            cnt = 0
        End If
        i = i + 1
    Loop
    arr(j) = Mid(text, i - cnt, cnt)
    n = j
End Sub


Public Sub sort(ByRef arr() As String, n As Integer)
    For i = 0 To n - 1
        For j = 0 To n - i - 1
            If Len(arr(j)) > Len(arr(j + 1)) Then
                Dim temp As String
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
    For i = 0 To n - 1
        For j = 0 To n - i - 1
            If Asc(arr(j)) > Asc(arr(j + 1)) Then
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
            End If
        Next j
    Next i
End Sub

Public Sub remduplication(ByRef done() As String, ByRef j As Integer, arr() As String, n As Integer)
    For i = 0 To n
        If i = 0 Then
            done(j) = arr(i)
            j = j + 1
        Else                                    ' a a b c c d e f f
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
        End If
    Next i
End Sub


