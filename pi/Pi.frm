VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Pi Calculator"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   Icon            =   "Pi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Pi.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   5580
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      Picture         =   "Pi.frx":0614
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Timer tmrTime 
      Interval        =   1
      Left            =   4920
      Top             =   5160
   End
   Begin VB.TextBox OutputBox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3000
      Width           =   7335
   End
   Begin VB.TextBox TextBox_LengthOfNumbers 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Text            =   "10"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton CalculateButton 
      Caption         =   "Calculate Pi !"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Pi The Never Ending Number!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   7335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Min."
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   4920
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Calculate Pi!"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblDTC 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Digits to calculate:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   765
      Index           =   2
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    Dim CalculatingPi As Integer  ' toggle true/false whether calc'ing pi
    
    '
    '   Infinite Sums Formulas:
    '
    '       Pi = 1/1 - 1/3 + 1/5 - 1/7 + 1/9 - 1/11 . . . = 4 / Pi
    '
    '       Pi = 1/1^2 + 1/2^2 + 1/3^2 +1/4^2 + 1/5^2 . . . = (Pi^2) / 6
    '
    '
    '   ArcTangent Formulas:
    '
    '       Pi = 4 * Atn(1)
    '
    '       Euler's Formula:
    '           Pi = 20 * Atn(1/7) + 8 * Atn(3/79)
    '
    '       Gauss's Formula:
    '           Pi = 48 * Atn(1/18) + 32 * Atn(1/57) - 20 * Atn(1/239)
    '
    '       Machin's Formula:
    '           Pi = 16 * Atn(1/5) - 4 * Atn(1/239)
    '
    '
    '       Power Series Expansion for ArcTangent:
    '           Atn(X) = X - X^3 /3 + X^5 /5 - X^7 /7 + X^9 /9 . . .
    '
    '
    '
    '   Ramanujan's Formulas:
    '
    '            1          1103   27493  1  1*3   53883  1*3  1*3*5*7
    '       -----------  =  ---- + -----  -  --- + -----  ---  ------- + . . .
    '       2*pi*Sqr(2)     99^2   99^6   2  4^2   99^10  2*4  4^2+8^2
    '
    '
    '       Elliptic Integral Formula:
    '
    '           1/pi = [ sqrt(8) / 9801 ] * sum { (4n)! * (1103+26390n) /
    '                  [(n!)^4 * 396^(4n) ] }      (n=0,1,2,... )
    

Sub CalculateButton_Click()

    If CalculatingPi = False Then
        CalculatePi
    Else
        End
    End If

End Sub

Sub CalculatePi()
    
    
    Dim TimeSpent As Double
    TimeSpent = Timer
    
    OutputBox = "Initializing": DoEvents
    CalculatingPi = True
    CalculateButton.Caption = "Stop!"

    Dim X As Integer
    Dim CarryPosition As Integer
    '  to be used in subtraction routine below
    
    Dim NumberOfLoops As Integer
    Dim LengthOfNumbers As Integer
    '  variables to be passed to FindArcTangent sub

    LengthOfNumbers = TextBox_LengthOfNumbers + 3
    '  add 3 extra places because last couple may not be accurate
    NumberOfLoops = Int(2 / 3 * LengthOfNumbers)
    '  each iteration should produce about 1 1/2 accurate places
    
    
    '  all numbers needed to be super accurate in this program
    '  are represented by arrays consisting of single character
    '  length strings.  the 1 position contains the digit in the
    '  number to the far left, and the >1 positions in the array
    '  represent the numbers going to the right in the # from there
    
    ReDim ArcTangent5(1 To LengthOfNumbers) As String * 1
    ReDim ArcTangent239(1 To LengthOfNumbers) As String * 1
    '  arrays to be calculated by FindArcTangent sub
    
    ReDim MultipliedArcTangent5(1 To LengthOfNumbers + 1) As String * 1
    ReDim MultipliedArcTangent239(1 To LengthOfNumbers + 1) As String * 1
    '  arrays to be calculated by MultiplyArray sub
    
    
    
    '       Machin's Formula:
    '           Pi = 16 * Atn(1/5) - 4 * Atn(1/239)

    OutputBox = "Calculating ArcTangent of 1/5": DoEvents
    FindArcTangent 5, NumberOfLoops, LengthOfNumbers, ArcTangent5()
    
    OutputBox = "Calculating the ArcTangent of 1/239": DoEvents
    FindArcTangent 239, NumberOfLoops, LengthOfNumbers, ArcTangent239()
    
    
    OutputBox = "Multiplying ArcTan of 1/5 by 16": DoEvents
    MultiplyArray ArcTangent5(), 16, MultipliedArcTangent5()

    OutputBox = "Multiplying ArcTan of 1/239 by 4": DoEvents
    MultiplyArray ArcTangent239(), 4, MultipliedArcTangent239()

    
    OutputBox = "Subtracting the Multiplied Arctangents": DoEvents
    For X = LengthOfNumbers To 1 Step -1
                      ' subtract MultipliedArcTangent239 array
                      ' from MultipliedArcTangent5 array
        If MultipliedArcTangent5(X) < MultipliedArcTangent239(X) Then
                                                '  do we need to carry?
            CarryPosition = X - 1 ' start with 1st number to the left
                  
            Do Until MultipliedArcTangent5(CarryPosition) <> "0"
                          ' find a non-zero number to borrow from
                MultipliedArcTangent5(CarryPosition) = "9"  'fill the other #'s
                CarryPosition = CarryPosition - 1         ' with 9's
                                ' go to the next number to the left
            Loop   '  loop until finding a non-zero number
             ' at end of loop, CarryPosition will be # to borrow from
            MultipliedArcTangent5(CarryPosition) = CStr(CInt(MultipliedArcTangent5(CarryPosition)) - 1)
                            ' decrease number carried from by one
            MultipliedArcTangent5(X) = CStr((CInt(MultipliedArcTangent5(X)) + 10) - CInt(MultipliedArcTangent239(X)))
          'add an extra ten (borrowed) to MultipliedArcTangent5 and subtract MultipliedArcTangent239
        Else ' just simple subtraction if there isn't carrying
        
            MultipliedArcTangent5(X) = CStr(CInt(MultipliedArcTangent5(X)) - CInt(MultipliedArcTangent239(X)))
           
        End If

    DoEvents
    Next X  ' loop to subtract entire MultipliedArcTangent239 array
        
    

    '  with the MultipliedArcTangent239 array subtracted from the
    '  MultipliedArcTangent5 array, the MultipliedArcTangent5 array
    '  should now be equal to pi


    Dim PiValue As String
    
    Label(2) = "Pi = 3. + . . .": DoEvents
    OutputBox = ""  ' clear text box
    For X = 1 To LengthOfNumbers - 3  ' don't print the extra 3 numbers
    '  dump the value of pi into the text box
    '  the array does not include the "3."
    '  the 3 was bumped out of the array in
    '  the multiplication routine
        
        PiValue = PiValue & MultipliedArcTangent5(X)
        If X Mod 5 = 0 Then
        '  insert a space every 50 places for word wrapping
            PiValue = PiValue & " "
        End If
    
    Next X

    OutputBox = PiValue

    
    MsgBox "Pi calculated to " & LengthOfNumbers - 3 & " decimal places." & Chr$(13) & "Completed " & NumberOfLoops & " iterations." & Chr$(13) & "Spent " & (Timer - TimeSpent) / 60 & " minutes calculating.", 64, "Calculations Complete"
    CalculatingPi = False
    CalculateButton.Caption = "Calculate Pi !"

End Sub

'                   Received                 Received                  Received                    Calculated and Passed
Sub FindArcTangent(ArcTanToFind As Integer, NumberOfLoops As Integer, LengthOfNumbers As Integer, ArcTangent() As String * 1)
    
    '  ArcTanToFind      reciprocal of number to find arctangent of
    '  NumberOfLoops     set number of iterations
    '  LengthOfNumbers   set length of numbers
    '
    '  Machin's Formula
    '  Pi = 16 * Atn(1/5) - 4 * Atn(1/239)
    '
    '  Atn(X) = X - X^3 /3 + X^5 /5 - X^7 /7 + X^9 /9 . . .
    
    
    Dim StartPos As Integer ' position to start division loops
    Dim Sum As Long   ' keeps track of total and carrying in adding loops
    Dim X As Integer  ' multiusage as counter in For...Next and Do loops
    Dim Divisor As Long  ' keeps track of what the Answer is to be divided by
    Dim Remainder As Long  ' remainder in the dividing loops
    Dim CarryPosition As Long  ' keeps track of position when carrying
    Dim DividedInto As Integer ' counts how many times # has divided into
    ReDim Answer(1 To LengthOfNumbers) As String * 1
    '  answer after being raised to a certain power, built on each loop
    ReDim Divided(1 To LengthOfNumbers) As String * 1
    '  the Answer after being divided by the divisor
    
    
    StartPos = 1
    
    For X = 1 To LengthOfNumbers
        ArcTangent(X) = "0"    '  change arrays from having
        Divided(X) = "0"       '  nulls to having 0's
        Answer(X) = "0"
    Next X

    
    Select Case ArcTanToFind
        Case 5
            ArcTangent(1) = "2"      '  final answer is .2 (1/5) so far
        
        Case 239
            X = 1
FillInNumbers:
            If X <= LengthOfNumbers Then ArcTangent(X) = "0": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "0": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "4": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "1": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "8": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "4": X = X + 1
            If X <= LengthOfNumbers Then ArcTangent(X) = "1": X = X + 1
                '  final answer is .0041841 repeating (1/239) so far
            If X <= LengthOfNumbers Then GoTo FillInNumbers
                '  fill in entire array with the repeating fraction
    End Select
    
    
    For X = 1 To LengthOfNumbers     '  answer will be the same as
        Answer(X) = ArcTangent(X)    '  the final arctangent at this point
    Next X
    
    
    
    Divisor = 3              '  start with the divisor being 3
    
    
    Do Until (Divisor - 1) / 2 = NumberOfLoops + 1 '  stops after formula
                                '  has been computed NumberOfLoops times
    
        For X = Int(StartPos) To LengthOfNumbers
                                '  loop to divide Answer array by #^2
            Remainder = Remainder * 10 ' multiply by ten and add new number
            Remainder = Remainder + CInt(Answer(X)) '  like bringing down
                                       ' the next number in long division
            Do Until Remainder < (ArcTanToFind ^ 2) ' loop until # is smaller
                Remainder = Remainder - (ArcTanToFind ^ 2) 'subtract and count
                DividedInto = DividedInto + 1 ' times it has gone into the #
            Loop

            Answer(X) = CStr(DividedInto)  ' the answer of the long division
            Divided(X) = Answer(X)    ' make a copy in the divided array
            DividedInto = 0    ' clear for next loop
    
            DoEvents
        Next X        '  loop for whole array

    
        DoneDividing = 0  ' reset this for next iteration
        Remainder = 0     ' clear variables for the next loop
        DividedInto = 0
    
    
        For X = Int(StartPos) To LengthOfNumbers
                                'loop to divide Divided array by Divisor
            Remainder = Remainder * 10       '  same long division loop
            Remainder = Remainder + CInt(Divided(X)) ' bring down number

            Do Until Remainder < Divisor        ' divide into remainder
                Remainder = Remainder - Divisor
                DividedInto = DividedInto + 1   ' count number of times
            Loop

            Divided(X) = CStr(DividedInto)  '  put answer back into array
            DividedInto = 0      ' clear variable for next loop
    
            DoEvents
        Next X     '  do this for entire Divided array

        Remainder = 0     ' clear variables for the next loop
        DividedInto = 0
        
        
        If Divisor Mod 4 = 1 Then ' all answers to be added will be true
            
            For X = LengthOfNumbers To 1 Step -1
                                 '  add Divided array to ArcTangent array
                Sum = Sum + CInt(Divided(X)) + CInt(ArcTangent(X))
                                             '  add the two numbers together
                ArcTangent(X) = CStr(Sum Mod 10)
                                 '  the answer will just be the ones' place
                Sum = Int(Sum / 10) '  divide the remainder by ten for
                     '  the increasing place value and drop the ones' place
                DoEvents
            Next X  '  loop for entire arrays
        
            Sum = 0  ' clear variable
        
        Else '  all answers to be subtracted will be false
            
            For X = LengthOfNumbers To 1 Step -1
                              ' subtract Divided array from ArcTan array
                If ArcTangent(X) < Divided(X) Then '  do we need to carry?
                
                    CarryPosition = X - 1 ' start with 1st number to the left
                    
                    Do Until ArcTangent(CarryPosition) <> "0"
                                  ' find a non-zero number to borrow from
                        ArcTangent(CarryPosition) = "9"  'fill the other #'s
                        CarryPosition = CarryPosition - 1         ' with 9's
                                        ' go to the next number to the left
                    Loop   '  loop until finding a non-zero number
                     ' at end of loop, CarryPosition will be # to borrow from
                    ArcTangent(CarryPosition) = CStr(CInt(ArcTangent(CarryPosition)) - 1)
                                    ' decrease number carried from by one
                    ArcTangent(X) = CStr((CInt(ArcTangent(X)) + 10) - CInt(Divided(X)))
                  'add an extra ten (borrowed) to ArcTan and subtract Divided
                Else ' just simple subtraction if there isn't carrying
            
                    ArcTangent(X) = CStr(CInt(ArcTangent(X)) - CInt(Divided(X)))
            
                End If

                DoEvents
            Next X  ' loop to subtract entire Divided array
        
            CarryPosition = 0  '  clear variable
        
        End If


        Divisor = Divisor + 2   ' each loop, power and divisor increase by 2
    
        OutputBox = "Calculating ArcTangent of 1/" & ArcTanToFind & ", Done with iteration " & (Divisor - 1) / 2
        DoEvents
    
        StartPos = StartPos + 1.25
    
    Loop  '  loop NumberOfLoops times
          '  each time ArcTangent gets more accurate


End Sub

'                  Received                         Received                       Calculated and Passed
Sub MultiplyArray(ArrayToMultiply() As String * 1, NumberToMultiplyBy As Integer, Answer() As String * 1)

    Dim Position As Integer  '  current position in array
    Dim SmallAnswer As Integer  '  keeps track of "sub-answers" in the multiplication process
    Dim NumberToCarry As Integer  '  keeps track of carrying
    
    For Position = TextBox_LengthOfNumbers + 3 To 1 Step -1
        
        SmallAnswer = (CInt(ArrayToMultiply(Position)) * NumberToMultiplyBy) + NumberToCarry
        '  multiply the 2 numbers together and add the remainder
        
        Answer(Position) = Right$(CStr(SmallAnswer), 1)
        '  add ones place of SmallAnswer to the whole answer

        If SmallAnswer < 10 Then  '  if greater than ten we will need
            NumberToCarry = 0     '  to carry
        Else
            NumberToCarry = CInt(Left$(CStr(SmallAnswer), CInt(Len(CStr(SmallAnswer))) - 1))
        End If
        '  carry the Answer without the ones place
        '  (everything is shifted to the right so it get divided by 10)

    
    DoEvents
    Next Position  ' go on to the next position (moving to the left)


End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label4_Click()
Me.WindowState = 1
End Sub

Private Sub tmrTime_Timer()
lblTime = Time
End Sub
