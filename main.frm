VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "TMDAI"
   ClientHeight    =   8985
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   449.25
   ScaleMode       =   2  '點
   ScaleWidth      =   452.25
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   12000
      TabIndex        =   2
      Top             =   2760
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1332
      Left            =   12000
      TabIndex        =   0
      Top             =   600
      Width           =   3372
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "main.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "0"
      Height          =   252
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   132
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'一空三的時候會有錯誤偵測雙活三
Dim black As Integer
Dim white As Integer
Dim turn As Boolean 'true黑 flase白
Dim bandw(16, 16) As Integer '黑1 白2 空0
Dim aibandw(16, 16) As Integer
Dim aiNum(15, 15) 'ai優先順序
Dim a1, a2  'AI掃描結果 a1四子 a3是活三死四
Dim space(111, 1) As Integer
Dim tick As Integer
Dim ctrl As Boolean

Dim impot As Integer






Private Sub Command1_Click()

Call aiCopy
ctrl = False

If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan5(x, y) = True Then
                Call drop(x, y)
                aibandw(x, y) = 0
                GoTo ed
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x

turn = Not (turn)
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan5(x, y) = True Then
                turn = Not (turn)
                Call drop(x, y)
                aibandw(x, y) = 0
                GoTo ed
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x

turn = Not (turn)
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan4(x, y) = True Then
                Call drop(x, y)
                aibandw(x, y) = 0
                GoTo ed
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x



Dim df4(10, 1) As Integer
Dim df4c As Integer
df4c = 0
turn = Not (turn)
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan4(x, y) = True Then
                df4(df4c, 0) = x
                df4(df4c, 1) = y
                df4c = df4c + 1
                aibandw(x, y) = 0
                
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x

For x = 0 To 111
    For y = 0 To 10
        If df4(y, 0) = space(x, 0) And df4(y, 1) = space(x, 1) Then
            turn = Not (turn)
            Call drop(space(x, 0), space(x, 1))
            GoTo ed
        End If
    Next y
Next
If df4(0, 0) <> 0 Then
    turn = Not (turn)
    Call drop(df4(0, 0), df4(0, 1))
    GoTo ed
End If


turn = Not (turn)
Dim choice(3, 2) As Integer
chcut = 0
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan3 >= 2 Then
                choice(0, 0) = x
                choice(0, 1) = y
                choice(0, 2) = impot
                aibandw(x, y) = 0
                GoTo md3
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x

md3:

turn = Not (turn)
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan3 >= 2 Then
                choice(1, 0) = x
                choice(1, 1) = y
                choice(1, 2) = impot
                aibandw(x, y) = 0
                GoTo md4
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x
md4:


turn = Not (turn)
If choice(1, 2) < choice(0, 2) And choice(1, 0) <> 0 Then
    Call drop(choice(1, 0), choice(1, 1))
    GoTo ed
ElseIf choice(1, 2) = choice(0, 2) And choice(0, 0) = 0 And choice(1, 0) <> 0 Then
    Call drop(choice(1, 0), choice(1, 1))
    GoTo ed
ElseIf choice(0, 2) <= choice(1, 2) And choice(0, 0) <> 0 Then
    Call drop(choice(0, 0), choice(0, 1))
    GoTo ed
End If

If tick < 30 Then GoTo jp


If turn = True Then
    bw = 1
Else
    bw = 2
End If


Dim df3(50, 1) As Integer
df3c = 0
For x = 1 To 15
    For y = 1 To 15
    
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan3 = 1 Then
                df3(df3c, 0) = x
                df3(df3c, 1) = y
                aibandw(x, y) = 0
                df3c = df3c + 1
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x
For x = 0 To 111
    For y = 0 To 50
        If df3(y, 0) = space(x, 0) And df3(y, 1) = space(x, 1) Then
            Call drop(space(x, 0), space(x, 1))
            GoTo ed
        End If
    Next y
Next
If df3(0, 0) <> 0 Then
    Call drop(df3(0, 0), df3(0, 1))
    GoTo ed
End If


Dim df32(50, 1) As Integer
df3c = 0
turn = Not (turn)
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For x = 1 To 15
    For y = 1 To 15
    
        If aibandw(x, y) = 0 Then
            aibandw(x, y) = bw
            If aiScan3s = 1 Then
                df32(df3c, 0) = x
                df32(df3c, 1) = y
                aibandw(x, y) = 0
                df3c = df3c + 1
            Else
                aibandw(x, y) = 0
            End If
        End If
    Next y
Next x
For x = 0 To 111
    For y = 0 To 50
        If df32(y, 0) = space(x, 0) And df32(y, 1) = space(x, 1) Then
        turn = Not (turn)
            Call drop(space(x, 0), space(x, 1))
            GoTo ed
        End If
    Next y
Next
If df32(0, 0) <> 0 Then
    turn = Not (turn)
    Call drop(df32(0, 0), df32(0, 1))
    GoTo ed
End If
turn = Not (turn)

jp:
If turn = True Then
    bw = 1
Else
    bw = 2
End If
For x = 0 To 111
    If space(x, 0) <> -1 Then
        If bandw(space(x, 0), space(x, 1)) = 0 Then
            Call drop(space(x, 0), space(x, 1))
            GoTo ed
        End If
    End If
Next
ed:
ctrl = True
End Sub
Private Sub aiCopy()
For x = 1 To 15
    For y = 1 To 15
        aibandw(x, y) = bandw(x, y)
    Next y
Next x
End Sub


Private Sub Command2_Click()
Call aiCopy
MsgBox aiScan3
End Sub

Private Sub Form_Activate()

ctrl = True
Open App.Path + "\space.txt" For Input As #1
For x = 0 To 111
    Input #1, space(x, 0)
    Input #1, space(x, 1)
Next
Close #1

Randomize
'For x = 0 To 14  '畫格子
'    Line (20, x * 30 + 20)-(440, x * 30 + 20)
'Next
'For y = 0 To 14
'    Line (y * 30 + 20, 20)-(y * 30 + 20, 440)
'Next
'
'
'For a = 0 To 15     '設定周圍格子
'    bandw(a, 0) = 3
'    bandw(16, a) = 3
'    bandw(a + 1, 16) = 3
'    bandw(0, a + 1) = 3
'Next

For a = 0 To 15     '設定周圍格子
    aibandw(a, 0) = 3
    aibandw(16, a) = 3
    aibandw(a + 1, 16) = 3
    aibandw(0, a + 1) = 3
Next

black = 0
white = 0
turn = True
FillStyle = 0
 Call drop(8, 8)
End Sub

Function aiScan4(i, j) As Boolean

aiScan4 = False
If turn = True Then
    bw = 1
Else
    bw = 2
End If

d1 = 0
d2 = 0
d3 = 0
d4 = 0
d5 = 0
d6 = 0
d7 = 0
d8 = 0

n = 1
While aibandw(i, j - n) = aibandw(i, j - 1) And aibandw(i, j - n) <> 0 And aibandw(i, j - n) = bw
    d1 = d1 + 1
    n = n + 1
Wend
If aibandw(i, j - n) = 0 Then
    d1 = d1 + 0.25
    n = n + 1
    If aibandw(i, j - n) = 0 Then
        d1 = d1 + 0.25
    End If
End If


n = 1
While aibandw(i + n, j - n) = aibandw(i + 1, j - 1) And aibandw(i + n, j - n) <> 0 And aibandw(i + n, j - n) = bw
    d2 = d2 + 1
    n = n + 1
Wend
If aibandw(i + n, j - n) = 0 Then
    d2 = d2 + 0.25
    n = n + 1
    If aibandw(i + n, j - n) = 0 Then
        d2 = d2 + 0.25
    End If
End If

n = 1
While aibandw(i + n, j) = aibandw(i + 1, j) And aibandw(i + n, j) <> 0 And aibandw(i + n, j) = bw
    d3 = d3 + 1
    n = n + 1
Wend
If aibandw(i + n, j) = 0 Then
    d3 = d3 + 0.25
    n = n + 1
    If aibandw(i + n, j) = 0 Then
        d3 = d3 + 0.25
    End If
End If

n = 1
While aibandw(i + n, j + n) = aibandw(i + 1, j + 1) And aibandw(i + n, j + n) <> 0 And aibandw(i + n, j + n) = bw
    d4 = d4 + 1
    n = n + 1
Wend
If aibandw(i + n, j + n) = 0 Then
    d4 = d4 + 0.25
    n = n + 1
    If aibandw(i + n, j + n) = 0 Then
        d4 = d4 + 0.25
    End If
End If

n = 1
While aibandw(i, j + n) = aibandw(i, j + 1) And aibandw(i, j + n) <> 0 And aibandw(i, j + n) = bw
    d5 = d5 + 1
    n = n + 1
Wend
If aibandw(i, j + n) = 0 Then
    d5 = d5 + 0.25
    n = n + 1
    If aibandw(i, j + n) = 0 Then
        d5 = d5 + 0.25
    End If
End If

n = 1
While aibandw(i - n, j + n) = aibandw(i - 1, j + 1) And aibandw(i - n, j + n) <> 0 And aibandw(i - n, j + n) = bw
    d6 = d6 + 1
    n = n + 1
Wend
If aibandw(i - n, j + n) = 0 Then
    d6 = d6 + 0.25
    n = n + 1
    If aibandw(i - n, j + n) = 0 Then
        d6 = d6 + 0.25
    End If
End If

n = 1
While aibandw(i - n, j) = aibandw(i - 1, j) And aibandw(i - n, j) <> 0 And aibandw(i - n, j) = bw
    d7 = d7 + 1
    n = n + 1
Wend
If aibandw(i - n, j) = 0 Then
    d7 = d7 + 0.25
    n = n + 1
    If aibandw(i - n, j) = 0 Then
        d7 = d7 + 0.25
    End If
End If

n = 1
While aibandw(i - n, j - n) = aibandw(i - 1, j - 1) And aibandw(i - n, j - n) <> 0 And aibandw(i - n, j - n) = bw
    d8 = d8 + 1
    n = n + 1
Wend
If aibandw(i - n, j - n) = 0 Then
    d8 = d8 + 0.25
    n = n + 1
    If aibandw(i - n, j - n) = 0 Then
        d8 = d8 + 0.25
    End If
End If
If Int(d1) + Int(d5) = 3 And d1 > Int(d1) And d5 > Int(d5) Then aiScan4 = True
If Int(d2) + Int(d6) = 3 And d2 > Int(d2) And d6 > Int(d6) Then aiScan4 = True
If Int(d3) + Int(d7) = 3 And d3 > Int(d3) And d7 > Int(d7) Then aiScan4 = True
If Int(d4) + Int(d8) = 3 And d4 > Int(d4) And d8 > Int(d8) Then aiScan4 = True

End Function


Function aiScan3s() As Integer
a2 = 0
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For i = 1 To 15
    For j = 1 To 15
        If aibandw(i, j) = 0 Then
            d1 = 0
            d2 = 0
            d3 = 0
            d4 = 0
            d5 = 0
            d6 = 0
            d7 = 0
            d8 = 0
            
            n = 1
            While aibandw(i, j - n) = aibandw(i, j - 1) And aibandw(i, j - n) <> 0 And aibandw(i, j - n) = bw
                d1 = d1 + 1
                n = n + 1
            Wend
            If aibandw(i, j - n) = 0 Then
                d1 = d1 + 0.25
                n = n + 1
                If aibandw(i, j - n) = 0 Then
                    d1 = d1 + 0.25
                End If
            End If
            
            
            n = 1
            While aibandw(i + n, j - n) = aibandw(i + 1, j - 1) And aibandw(i + n, j - n) <> 0 And aibandw(i + n, j - n) = bw
                d2 = d2 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j - n) = 0 Then
                d2 = d2 + 0.25
                n = n + 1
                If aibandw(i + n, j - n) = 0 Then
                    d2 = d2 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i + n, j) = aibandw(i + 1, j) And aibandw(i + n, j) <> 0 And aibandw(i + n, j) = bw
                d3 = d3 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j) = 0 Then
                d3 = d3 + 0.25
                n = n + 1
                If aibandw(i + n, j) = 0 Then
                    d3 = d3 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i + n, j + n) = aibandw(i + 1, j + 1) And aibandw(i + n, j + n) <> 0 And aibandw(i + n, j + n) = bw
                d4 = d4 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j + n) = 0 Then
                d4 = d4 + 0.25
                n = n + 1
                If aibandw(i + n, j + n) = 0 Then
                    d4 = d4 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i, j + n) = aibandw(i, j + 1) And aibandw(i, j + n) <> 0 And aibandw(i, j + n) = bw
                d5 = d5 + 1
                n = n + 1
            Wend
            If aibandw(i, j + n) = 0 Then
                d5 = d5 + 0.25
                n = n + 1
                If aibandw(i, j + n) = 0 Then
                    d5 = d5 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j + n) = aibandw(i - 1, j + 1) And aibandw(i - n, j + n) <> 0 And aibandw(i - n, j + n) = bw
                d6 = d6 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j + n) = 0 Then
                d6 = d6 + 0.25
                n = n + 1
                If aibandw(i - n, j + n) = 0 Then
                    d6 = d6 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j) = aibandw(i - 1, j) And aibandw(i - n, j) <> 0 And aibandw(i - n, j) = bw
                d7 = d7 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j) = 0 Then
                d7 = d7 + 0.25
                n = n + 1
                If aibandw(i - n, j) = 0 Then
                    d7 = d7 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j - n) = aibandw(i - 1, j - 1) And aibandw(i - n, j - n) <> 0 And aibandw(i - n, j - n) = bw
                d8 = d8 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j - n) = 0 Then
                d8 = d8 + 0.25
                n = n + 1
                If aibandw(i - n, j - n) = 0 Then
                    d8 = d8 + 0.25
                End If
            End If
            
            If Int(d1) = 3 And Int(d1) < d1 And d1 + d5 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d2) = 3 And Int(d2) < d2 And d2 + d6 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d3) = 3 And Int(d3) < d3 And d3 + d7 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d4) = 3 And Int(d4) < d4 And d4 + d8 >= 3.5 Then
                a2 = a2 + 1
            End If
                 
            If Int(d1) + Int(d5) >= 4 Then a2 = a2 + 1
            If Int(d2) + Int(d6) >= 4 Then a2 = a2 + 1
            If Int(d3) + Int(d7) >= 4 Then a2 = a2 + 1
            If Int(d4) + Int(d8) >= 4 Then a2 = a2 + 1
            
        End If
    Next j
Next i
aiScan3s = a2
End Function

Function aiScan3() As Integer
a2 = 0
impot = False
If turn = True Then
    bw = 1
Else
    bw = 2
End If

For i = 1 To 15
    For j = 1 To 15
        If aibandw(i, j) = 0 Then
            d1 = 0
            d2 = 0
            d3 = 0
            d4 = 0
            d5 = 0
            d6 = 0
            d7 = 0
            d8 = 0
            
            n = 1
            While aibandw(i, j - n) = aibandw(i, j - 1) And aibandw(i, j - n) <> 0 And aibandw(i, j - n) = bw
                d1 = d1 + 1
                n = n + 1
            Wend
            If aibandw(i, j - n) = 0 Then
                d1 = d1 + 0.25
                n = n + 1
                If aibandw(i, j - n) = 0 Then
                    d1 = d1 + 0.25
                End If
            End If
            
            
            n = 1
            While aibandw(i + n, j - n) = aibandw(i + 1, j - 1) And aibandw(i + n, j - n) <> 0 And aibandw(i + n, j - n) = bw
                d2 = d2 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j - n) = 0 Then
                d2 = d2 + 0.25
                n = n + 1
                If aibandw(i + n, j - n) = 0 Then
                    d2 = d2 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i + n, j) = aibandw(i + 1, j) And aibandw(i + n, j) <> 0 And aibandw(i + n, j) = bw
                d3 = d3 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j) = 0 Then
                d3 = d3 + 0.25
                n = n + 1
                If aibandw(i + n, j) = 0 Then
                    d3 = d3 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i + n, j + n) = aibandw(i + 1, j + 1) And aibandw(i + n, j + n) <> 0 And aibandw(i + n, j + n) = bw
                d4 = d4 + 1
                n = n + 1
            Wend
            If aibandw(i + n, j + n) = 0 Then
                d4 = d4 + 0.25
                n = n + 1
                If aibandw(i + n, j + n) = 0 Then
                    d4 = d4 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i, j + n) = aibandw(i, j + 1) And aibandw(i, j + n) <> 0 And aibandw(i, j + n) = bw
                d5 = d5 + 1
                n = n + 1
            Wend
            If aibandw(i, j + n) = 0 Then
                d5 = d5 + 0.25
                n = n + 1
                If aibandw(i, j + n) = 0 Then
                    d5 = d5 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j + n) = aibandw(i - 1, j + 1) And aibandw(i - n, j + n) <> 0 And aibandw(i - n, j + n) = bw
                d6 = d6 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j + n) = 0 Then
                d6 = d6 + 0.25
                n = n + 1
                If aibandw(i - n, j + n) = 0 Then
                    d6 = d6 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j) = aibandw(i - 1, j) And aibandw(i - n, j) <> 0 And aibandw(i - n, j) = bw
                d7 = d7 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j) = 0 Then
                d7 = d7 + 0.25
                n = n + 1
                If aibandw(i - n, j) = 0 Then
                    d7 = d7 + 0.25
                End If
            End If
            
            n = 1
            While aibandw(i - n, j - n) = aibandw(i - 1, j - 1) And aibandw(i - n, j - n) <> 0 And aibandw(i - n, j - n) = bw
                d8 = d8 + 1
                n = n + 1
            Wend
            If aibandw(i - n, j - n) = 0 Then
                d8 = d8 + 0.25
                n = n + 1
                If aibandw(i - n, j - n) = 0 Then
                    d8 = d8 + 0.25
                End If
            End If
            
            If Int(d1) = 3 And Int(d1) < d1 And d1 + d5 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d2) = 3 And Int(d2) < d2 And d2 + d6 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d3) = 3 And Int(d3) < d3 And d3 + d7 >= 3.5 Then
                a2 = a2 + 1
            End If
            If Int(d4) = 3 And Int(d4) < d4 And d4 + d8 >= 3.5 Then
                a2 = a2 + 1
            End If
            
            If ((Int(d1) = 2 And Int(d5) = 1) Or (Int(d1) = 1 And Int(d5) = 2)) And Int(d1) < d1 And Int(d5) < d5 Then a2 = a2 + 1
            If ((Int(d2) = 2 And Int(d6) = 1) Or (Int(d2) = 1 And Int(d6) = 2)) And Int(d2) < d2 And Int(d6) < d6 Then a2 = a2 + 1
            If ((Int(d3) = 2 And Int(d7) = 1) Or (Int(d3) = 1 And Int(d7) = 2)) And Int(d3) < d3 And Int(d7) < d7 Then a2 = a2 + 1
            If ((Int(d4) = 2 And Int(d8) = 1) Or (Int(d4) = 1 And Int(d8) = 2)) And Int(d4) < d4 And Int(d8) < d8 Then a2 = a2 + 1
            
            If Int(d1) + Int(d5) >= 4 Then
                a2 = a2 + 1
                impot = True
            End If
            If Int(d2) + Int(d6) >= 4 Then
                a2 = a2 + 1
                impot = True
            End If
            If Int(d3) + Int(d7) >= 4 Then
                a2 = a2 + 1
                impot = True
            End If
            If Int(d4) + Int(d8) >= 4 Then
                a2 = a2 + 1
                impot = True
            End If
            
        End If
    Next j
Next i
aiScan3 = a2
End Function

Function aiScan5(x, y) As Boolean
aiScan5 = False
If turn = True Then
    bw = 1
Else
    bw = 2
End If

a = 0
Do
    If aibandw(x + a, y) = bw Then
        a = a + 1
    Else
        Exit Do
    End If
Loop
b = 0
Do
    If aibandw(x - b, y) = bw Then
        b = b + 1
    Else
        Exit Do
    End If
Loop
If a + b >= 6 Then aiScan5 = True


a = 0
Do
    If aibandw(x, y + a) = bw Then
        a = a + 1
    Else
        Exit Do
    End If
Loop
b = 0
Do
    If aibandw(x, y - b) = bw Then
        b = b + 1
    Else
        Exit Do
    End If
Loop
If a + b >= 6 Then aiScan5 = True

a = 0
Do
    If aibandw(x + a, y + a) = bw Then
        a = a + 1
    Else
        Exit Do
    End If
Loop
b = 0
Do
    If aibandw(x - b, y - b) = bw Then
        b = b + 1
    Else
        Exit Do
    End If
Loop
If a + b >= 6 Then aiScan5 = True

a = 0
Do
    If aibandw(x + a, y - a) = bw Then
        a = a + 1
    Else
        Exit Do
    End If
Loop
b = 0
Do
    If aibandw(x - b, y + b) = bw Then
        b = b + 1
    Else
        Exit Do
    End If
Loop
If a + b >= 6 Then aiScan5 = True
End Function


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim posx As Integer
Dim posy As Integer

posx = (x - 345) / 600 + 1
posy = (y - 345) / 600 + 1
If posx < 16 And posx > 0 And posy > 0 And posy < 16 And bandw(posx, posy) = 0 And ctrl = True Then
    Call drop(posx, posy)
    Call Command1_Click
End If

End Sub


Private Sub drop(ByVal posx As Integer, ByVal posy As Integer)

If bandw(posx, posy) = 0 Then
    If turn = True Then
        FillColor = RGB(0, 0, 0)
        Circle (30 * posx - 10, 30 * posy - 10), 14
        bandw(posx, posy) = 1
        turn = Not (turn)
    Else
        FillColor = RGB(255, 255, 255)
        Circle (30 * posx - 10, 30 * posy - 10), 14
        bandw(posx, posy) = 2
        turn = Not (turn)
    End If
    For x = 0 To 111
        If posx = space(x, 0) And posy = space(x, 1) Then
            space(x, 0) = -1
            space(x, 1) = -1
        End If
    Next
    Call chkwin(posx, posy)
End If
End Sub



Private Sub chkwin(ByVal posx As Integer, ByVal posy As Integer)
If turn = True Then
    bw = 2
Else
    bw = 1
End If
If bw = 2 Then
    win = "黑棋"
Else
    win = "白棋"
End If

a = 1
While bandw(posx, posy + a) = bw
    a = a + 1
Wend
b = 1
While bandw(posx, posy - b) = bw
    b = b + 1
Wend
If a + b >= 6 Then
    MsgBox (win + "勝利")
    End
End If

a = 1
While bandw(posx + a, posy) = bw
    a = a + 1
Wend
b = 1
While bandw(posx - b, posy) = bw
    b = b + 1
Wend
If a + b >= 6 Then
    MsgBox (win + "勝利")
    End
End If

a = 1
While bandw(posx + a, posy + a) = bw
    a = a + 1
Wend
b = 1
While bandw(posx - b, posy - b) = bw
    b = b + 1
Wend

If a + b >= 6 Then
    MsgBox (win + "勝利")
    End
End If

a = 1
While bandw(posx + a, posy - a) = bw
    a = a + 1
Wend
b = 1
While bandw(posx - b, posy + b) = bw
    b = b + 1
Wend
If a + b >= 6 Then
    MsgBox (win + "勝利")
    End
End If


End Sub
