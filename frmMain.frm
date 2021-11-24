VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdGetIncidents 
      Caption         =   "Start"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdUniqueLots 
      Caption         =   "Unique Lots"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcLotNumber 
      Caption         =   "Process Lot Number"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdFilterLotNumber 
      Caption         =   "Make Lot#"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cdmGetData 
      Caption         =   "Get Data"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Caption         =   "..."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Get only "COVID19" entries
Private Sub cdmGetData_Click()
    
    Dim strLine As String
    Dim intCnt As Long
    
    Open App.Path & "\vax.csv" For Input As #1
    Open App.Path & "\vax2.csv" For Output As #2
    
    While Not EOF(1)
      Line Input #1, strLine
      If GetInfoFromString(strLine, 2) = "COVID19" Then
        Print #2, strLine
        intCnt = intCnt + 1
      End If
    Wend
    
    Close #1
    Close #2
    Debug.Print intCnt & " lines processed."
    
End Sub

'Make list of lot numbers
Private Sub cmdFilterLotNumber_Click()
    
    Dim strLine As String
    Dim intCnt As Long
    Dim strLNumber As String
    
    Open App.Path & "\vax2.csv" For Input As #1
    Open App.Path & "\lots.csv" For Output As #2
    
    intCnt = 1
    
    While Not EOF(1)
      Line Input #1, strLine
      strLNumber = GetInfoFromString(strLine, 4)
      If strLNumber <> "" And InStr(1, strLNumber, UCase("know")) = 0 Then
          Print #2, GetInfoFromString(strLine, 3) & "," & strLNumber
          intCnt = intCnt + 1
          lblMsg.Caption = Format(intCnt)
          lblMsg.Refresh
      End If
    Wend
    
    Close #1
    Close #2
    Debug.Print intCnt & " lines processed."
End Sub

Private Sub cmdGetIncidents_Click()

   Dim strLot(30000) As String
   Dim intEntry(30000) As Integer
   Dim T1 As Integer
   Dim intCnt As Long
   Dim strS As String
   Dim strLotID As String
   Dim intLine As Long
   
   intLine = 0
   intCnt = 1
   
   Open App.Path & "\lots3.csv" For Input As #1
   While Not EOF(1)
     Line Input #1, strS
     strLot(intCnt) = strS
     intCnt = intCnt + 1
   Wend
   Close #1
   
   Open App.Path & "\lots2.csv" For Input As #2
   While Not EOF(2)
     Line Input #2, strS
     intLine = intLine + 1
     strLotID = GetInfoFromString(strS, 2)
       
     For T1 = 1 To intCnt
       If strLotID = strLot(T1) Then
         intEntry(T1) = intEntry(T1) + 1
       End If
     Next
     lblMsg.Caption = Format(intLine)
     lblMsg.Refresh

   Wend
   Close #2
   
   Open App.Path & "\lots4.csv" For Output As #3
   
   For T1 = 1 To intCnt
     If Len(strLot(T1)) >= 3 And Len(strLot(T1)) <= 8 Then
       Print #3, strLot(T1) & "," & intEntry(T1)
     End If
   Next
   
   Close #3
   
   Debug.Print "Done."
   
End Sub

Private Sub cmdProcLotNumber_Click()

    Dim strLine As String
    Dim intCnt As Long
    Dim strLNumber As String
    Dim strCompany As String
    Dim T1 As Integer
    Dim intLine As Long
    Dim strOut As String
    
    Open App.Path & "\lots.csv" For Input As #1
    Open App.Path & "\lots2.csv" For Output As #2
    
    intCnt = 0
    
    While Not EOF(1)
      Line Input #1, strLine
      intLine = intLine + 1
      strCompany = GetInfoFromString(strLine, 1)
      strLNumber = UCase(GetInfoFromString(strLine, 2))
      If Left$(strCompany, 4) = "MODE" Then
        strOut = ""
        For T1 = 1 To Len(strLNumber)
          If Asc(Mid$(strLNumber, T1, 1)) > 32 Then
            If Asc(Mid$(strLNumber, T1, 1)) <> Asc("#") Then
              strOut = strOut & Mid$(strLNumber, T1, 1)
            End If
          End If
        Next
        Print #2, "MODERNA" & "," & strOut
      End If
      If Left$(strCompany, 4) = "PFIZ" Then
        strOut = ""
        For T1 = 1 To Len(strLNumber)
            If Asc(Mid$(strLNumber, T1, 1)) > 32 Then
                strOut = strOut & Mid$(strLNumber, T1, 1)
            End If
        Next
        Print #2, "PFIZER" & "," & Right$(strOut, 6)
      End If
      
      lblMsg.Caption = Format(intLine)
      lblMsg.Refresh
    Wend
    
    Close #1
    Close #2
    Debug.Print intCnt & " different lot numbers found."
End Sub

Private Sub cmdUniqueLots_Click()
    
    Dim strLNumber As String
    Dim strCompany As String
    Dim intCnt As Long
    Dim intLine As Long
    Dim strLine As String
    Dim intFound As Integer
    Dim T1 As Long
    
    Open App.Path & "\lots2.csv" For Input As #1
    Open App.Path & "\lots3.csv" For Output As #2
    
    intCnt = 1
    
    While Not EOF(1)
      Line Input #1, strLine
      intLine = intLine + 1
      strCompany = GetInfoFromString(strLine, 1)
      strLNumber = UCase(GetInfoFromString(strLine, 2))
      intFound = 0
      For T1 = 1 To intCnt
        If strLNumber = strData(T1) Then
            intFound = 1
            T1 = intCnt
        End If
      Next
      If intFound = 0 Then
        intCnt = intCnt + 1
        strData(intCnt) = strLNumber
      End If
      lblMsg.Caption = Format(intCnt) & " " & Format(intLine)
      lblMsg.Refresh

    Wend
    
    For T1 = 1 To intCnt
      Print #2, strData(T1)
    Next
    Close #1
    Close #2
      
    
End Sub

Private Sub Command1_Click()

    Debug.Print GetInfoFromString("0916602,COVID19,PFIZER\BIONTECH,EL1284,1,IM,LA,COVID19 (COVID19 (PFIZER-BIONTECH))", 1)
    Debug.Print GetInfoFromString("0916602,COVID19,PFIZER\BIONTECH,EL1284,1,IM,LA,COVID19 (COVID19 (PFIZER-BIONTECH))", 4)
    Debug.Print GetInfoFromString("0916602,COVID19,PFIZER\BIONTECH,EL1284,1,IM,LA,COVID19 (COVID19 (PFIZER-BIONTECH))", 8)
    
    
End Sub
