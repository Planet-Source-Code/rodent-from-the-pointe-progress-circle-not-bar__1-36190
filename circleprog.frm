VERSION 5.00
Begin VB.Form circeProg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "circular progress bar"
   ClientHeight    =   3255
   ClientLeft      =   2040
   ClientTop       =   2850
   ClientWidth     =   2895
   Icon            =   "circleprog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2895
   Begin VB.CommandButton Command1 
      Caption         =   "set percent"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox setPer 
      Height          =   285
      Left            =   480
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "35"
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   480
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Tag             =   "0"
      Top             =   480
      Width           =   1935
      Begin VB.Label perCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   210
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " %"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   165
   End
End
Attribute VB_Name = "circeProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function drawProgress(percentDone As Integer) As Integer
'this function draws and processes all the math of the progress circle

'if an unacceptable percent number is passed, give an error
If percentDone > 100 Then Err.Raise 5: Exit Function

'show the percent caption when function is started for the first time
perCap.Visible = True

'dimensionalize and set some variables of the canvas
Dim pWidth As Long: pWidth = (canvas.Width / 15) - 4
Dim pHeight As Long: pHeight = (canvas.Height / 15) - 4

'dimensionalize some variables for color
Dim enabledColor As Long: enabledColor = RGB(0, 255, 0)
Dim disabledcolor As Long: disabledcolor = RGB(255, 0, 0)

'dimensionalize some other variables for doing correct math
Dim w As Long
Dim eachPercent As Long

'reposition the caption to be offset center
perCap.Move (pWidth * 15) / 2, (pHeight * 15) / 2
perCap.Caption = percentDone & "%"

'the following for..next statements see how many different
'positions there are around the outside of the canvas
For i = 0 To pWidth
    w = w + 2
Next i
For i = 0 To pHeight
    w = w + 2
Next i

'calculate that every 1 percent is every w/100 along
'the edge of the canvas
eachPercent = Int(w / 100)

'set temporary variables and find the first 25 percent values
tempW = 0
tempH = 1
For i = 1 To 25
    If tempH < pHeight Then
    perPos(i).x = tempW
    perPos(i).y = tempH
    tempH = tempH + eachPercent
    End If
Next i

'set temporary variables and find the 26-50 percent value
tempW = 0
tempH = 0
For i = 26 To 50
    If tempW < pWidth Then
    perPos(i).x = tempW
    perPos(i).y = pHeight
    tempW = tempW + eachPercent
    End If
Next i

'set temporary variables and find the 51-75 percent values
tempW = pWidth
tempH = pHeight
For i = 51 To 75
    If tempH > 0 Then
    perPos(i).x = pWidth
    perPos(i).y = tempH
    tempH = tempH - eachPercent
    End If
Next i

'set temporary variables and find the 76-100 percent values
tempW = pWidth
tempH = 0
For i = 76 To 100
    If tempW > 0 Then
    perPos(i).x = tempW
    perPos(i).y = tempH
    tempW = tempW - eachPercent
    End If
Next i

'clear the canvas to ready for a new picture
canvas.Cls

'if the percent to calculate is either 0 or 100
'goto the special line number to handle that
If percentDone = 0 Then GoTo noper
If percentDone = 100 Then GoTo fullper

'otherwise, process the percent...
'make sure to set the fillstye to transparent so the ellipse doesn't fill with color
canvas.FillStyle = 1

'create the ellipse AKA circle
Ellipse canvas.hdc, 0, 0, pWidth, pHeight

'make sure the fillstyle is solid to correct fill the percent of the piegraph
'also, set the color
canvas.FillStyle = 0
canvas.FillColor = enabledColor

'make the piegraph
Pie canvas.hdc, 0, 0, pWidth, pHeight, 0, 0, perPos(percentDone).x, perPos(percentDone).y

'change the color to show the incomplete percent
canvas.FillColor = disabledcolor

'fill incomplete percent with color
ExtFloodFill canvas.hdc, (pWidth / 5) + 2, (pHeight / 5) - 2, canvas.BackColor, 1

'percent caption text
m = percentDone & "%"
TextOut canvas.hdc, pWidth / 2, pHeight / 2, m, Len(m)

'refresh the canvas
canvas.Refresh

'leave the function prematurely to skip over next part of code
Exit Function




'process 0 percent
noper:
'fill solid
canvas.FillStyle = 0
canvas.FillColor = RGB(255, 0, 0)

'draw ellipse
Ellipse canvas.hdc, 0, 0, pWidth, pHeight

'percent caption text
m = percentDone & "%"
TextOut canvas.hdc, pWidth / 2, pHeight / 2, m, Len(m)

'show picture
canvas.Refresh
Exit Function



'process 100 percent
fullper:
'fill transparent
canvas.FillStyle = 1

'draw circle
Ellipse canvas.hdc, 0, 0, pWidth, pHeight

'fill solid
canvas.FillStyle = 0

'change fill color
canvas.FillColor = enabledColor

'fill canvas point with color
ExtFloodFill canvas.hdc, (pWidth / 2), (pHeight / 2), canvas.BackColor, 1

m = percentDone & "%"
TextOut canvas.hdc, pWidth / 2, pHeight / 2, m, Len(m)

'show picture
canvas.Refresh

Exit Function


End Function


Private Sub Command1_Click()
drawProgress Val(setPer.Text)
End Sub


Private Sub Form_Load()
perCap.ForeColor = 0
End Sub


