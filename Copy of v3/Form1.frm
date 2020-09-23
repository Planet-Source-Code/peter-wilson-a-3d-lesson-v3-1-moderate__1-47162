VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "MIDAR's Simple 3D Lesson"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerZoomIn 
      Interval        =   1
      Left            =   720
      Top             =   60
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   60
      Top             =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API used for reading the keyboard.
' ==================================
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer


' Virtual Camera: World Position, LookAt point, Tilt, FOV, Zoom, Near & Far Clipping plane.
' This is pretty comprehensive stuff!
' =========================================================================================
Private m_Camera As mdrVector4
Private m_CameraFOV As Single               ' Field Of View (FOV). "90 degree FOV" = "1x Zoom". IMPORTANT: Adjusting either FOV and ZOOM both do essentially the *same* thing!
Private m_CameraZoom As Single              ' (The Zoom value is calculated from the FOV - see: 'UpdateCameraParameters')
Private m_CameraLookAt As mdrVector4
Private m_ClippingDistanceFar As Single     ' Dots far away are not drawn.
Private m_ClippingDistanceNear As Single    ' Dots that are too close, are not drawn.


' m_Dots & m_Temp will hold an Array of Vectors (as defined in 'mMaths' module.)
' =================================================================================
Private m_Dots() As mdrVector4  ' << We define our Dots only once, and store them here.
Private m_Temp() As mdrVector4  ' << The original dots are transformed (by the camera code) into this temporary storage area.

Public Function CalculateZoom(FOV As Single) As Single

    ' Given a Field Of View, calculate the Zoom.
    CalculateZoom = 1 / Tan(ConvertDeg2Rad(FOV) / 2)
    
End Function
Public Function CalculateFOV(Zoom As Single) As Single

    ' Given a Zoom value, calculate the 'Field Of View'
    CalculateFOV = ConvertRad2Deg(2 * Atn(1 / Zoom))
    
End Function

Public Sub DrawDots()

    ' ===============================
    ' Draws the Dots onto the screen.
    ' ===============================
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    Dim sngDistance As Single
    
    Dim intBrightness As Integer
    Dim sngDeltaVisible As Single   ' Distance between the near and far clip distances.
    sngDeltaVisible = m_ClippingDistanceFar - m_ClippingDistanceNear
    
    
    ' Set the drawing style and width, etc.
    ' =====================================
    Me.DrawWidth = 1                    '   << Set the Width of the Pen. Any value higher than 1 will slow down animation.
    
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Temp) To UBound(m_Temp)
                    
        PixelX = m_Temp(lngIndex).x
        PixelY = -m_Temp(lngIndex).y        ' Negated to make positive Y go up (and not down like Microsoft wants us to do... this is an aesthetics issue.)
        sngDistance = m_Temp(lngIndex).Z    ' The Z coordinate, now represents the distance between the current Dot and the Camera.
        
        ' Only draw dots in front of the camera (and not behind us),
        ' but no further away than the Far clipping distance.
        ' ==========================================================
        If (sngDistance > m_ClippingDistanceNear) And (sngDistance < m_ClippingDistanceFar) Then
        
            ' Ignore Pixels that extend outside of the viewing window. Although the OS will pretty much do this
            ' for us 99% of the time, it fails with 'OverFlow Errors' the rest of the time when the OS tries to
            ' plot extreamly large pixel values... I consider this a Microsoft bug. Good on ya MS! :-p
            If (Abs(PixelX) < (1 / m_CameraZoom)) And (Abs(PixelY) < (1 / m_CameraZoom)) Then
            
                ' Shading dots is an important depth-cue.
                intBrightness = 255 - CInt((sngDistance / sngDeltaVisible) * 255)
                
                ' Plot the point
                Me.PSet (PixelX, PixelY), RGB(intBrightness, intBrightness, intBrightness)
                
            End If ' Is Pixel within the window?
        End If ' Is the Pixel within the near & far clip values?
        
     Next lngIndex
    
    Exit Sub
errTrap:
    ' Just ignore any errors.
    
End Sub


Private Sub CreateTestData()

    ' =====================================================
    ' Create a nice big test grid (and some ground clutter)
    ' =====================================================
    
    Screen.MousePointer = vbHourglass
    
    Dim lngIndex As Long
    
    Dim intX As Integer
    Dim intY As Integer
    Dim intZ As Integer
    
    lngIndex = -1 ' Reset to -1, because we'll soon be increasing this value to 0 (the start of our array)
    
    ' ===============================================================
    ' Create some random ground clutter (ie. grass blades, whatever?)
    ' (If you are feeling adventurous, you might like to introduce
    ' colour into this application to make the grass green.
    ' ===============================================================
    For intX = 0 To 200                             '   << Try increase the number of ground clutter dots
        lngIndex = lngIndex + 1
        ReDim Preserve m_Dots(lngIndex)
        m_Dots(lngIndex).x = (Rnd * 100) - 50       '   << ie. Random number between -50 and +50
        m_Dots(lngIndex).y = 0                      '   << Because this is the ground, the elevation is zero.
        m_Dots(lngIndex).Z = (Rnd * 100) - 50
        m_Dots(lngIndex).W = 1
    Next intX
    
    
    ' ====================================================================
    ' Create 3 large lines out of dots, representing the 3 axes (x, y & z)
    ' ====================================================================
    For intX = -100 To 100 Step 2                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 2              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 2          '   << Positive Y goes Up
                
                If (intX = 0 And intY = 0) Or (intX = 0 And intZ = 0) Or (intY = 0 And intZ = 0) Then
                
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX
                    m_Dots(lngIndex).y = intY
                    m_Dots(lngIndex).Z = intZ
                    m_Dots(lngIndex).W = 1
                    
                End If
            Next intY
        Next intZ
    Next intX
    
    
    ' ===================
    ' Create a Test Grid.
    ' ===================
    For intX = -100 To 100 Step 5                  '   << Positive X points to the Right
        For intZ = -100 To 100 Step 5              '   << Positive Z points *into* the monitor - away from you.
            For intY = -100 To 100 Step 5          '   << Positive Y goes Up

                If (Abs(intX) = 100 Or Abs(intZ) = 100 Or intX = 0 Or intZ = 0) And (intY = 0 Or Abs(intY) = 100) Then

                    ' Create the basement (below ground level), floor and roof (of the test grid)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX / 100
                    m_Dots(lngIndex).y = intY / 100
                    m_Dots(lngIndex).Z = intZ / 100
                    m_Dots(lngIndex).W = 1
                    
                ElseIf Abs(intX) = 100 And Abs(intZ) = 100 Then

                    ' Put some corners on it (ie. 4 support beams)
                    lngIndex = lngIndex + 1
                    ReDim Preserve m_Dots(lngIndex)
                    m_Dots(lngIndex).x = intX / 100
                    m_Dots(lngIndex).y = intY / 100
                    m_Dots(lngIndex).Z = intZ / 100
                    m_Dots(lngIndex).W = 1
                    
                End If
            Next intY
        Next intZ
    Next intX
    
    For intX = -100 To 100 Step 10       '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
        For intZ = -100 To 100 Step 10   '   << Try adjusting the Step value to 1, 2, 5, 10 or 25.
            
            lngIndex = lngIndex + 1
            ReDim Preserve m_Dots(lngIndex)
            m_Dots(lngIndex).x = intX / 100                                 '   << Positive X points to the Right
            m_Dots(lngIndex).y = Cos(Sqr(intX * intX + intZ * intZ) / 30)   '   << Positive Y goes Up
            m_Dots(lngIndex).Z = intZ / 100                                 '   << Positive Z points *into* the monitor - away from you.
            m_Dots(lngIndex).W = 1

        Next intZ
    Next intX
    
           
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub DrawCrossHairs()

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Me.DrawWidth = 1
    
    ' Draw Horizontal line
    Me.ForeColor = RGB(0, 32, 32)
    Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
    
    ' Draw Vertical line
    Me.ForeColor = RGB(0, 64, 64)
    Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
    
End Sub

Private Sub DrawParameters(InDemo As Boolean, DemoCounter As Long)

    ' ==========================================================================================
    ' This routine slows down the program, because printing text is very slow.
    ' Speed has been sacrificed for instructional clarity for beginners to 3D Computer Graphics.
    ' Remember that by-and-large I am programming things the slow way, in an effort to be clear.
    ' You can always speed up my code by making your own clever adjustments.
    ' ==========================================================================================
    
    Dim sngX As Single
    
    ' Set our start printing position
    ' Remember, The origin of our screen has been moved into the center of the window, but we want text top-left.
    Me.ForeColor = RGB(255, 255, 192)
    sngX = Me.ScaleLeft
    Me.CurrentY = Me.ScaleTop
    
    
    ' Show product name.
    Me.CurrentX = sngX
    Me.Print App.ProductName & " - " & App.LegalCopyright
    
    
    ' Show helpful reminders.
    If InDemo = False Then
        Me.CurrentX = sngX
        Me.Print "Keys:  ESC, Left/Right/Up/Down, Shift-Up/Down, Page-Up/Down, Spacebar. Modify Camera LookAt point in code: 'm_CameraLookAt'"
        
        ' Show helpful reminders.
        Me.CurrentX = sngX
        Me.Print "Mouse:  Move mouse over dots to display original defined coordinates." & vbNewLine
    End If
    
    ' Show current Camera position.
    Me.CurrentX = sngX
    Me.Print "Camera:  x: " & Format(m_Camera.x, "Fixed") & "   y: " & Format(m_Camera.y, "Fixed") & "   z: " & Format(m_Camera.Z, "Fixed")
    
    ' Show current LookAt point.
    Me.CurrentX = sngX
    Me.Print "LookAt:  x: " & Format(m_CameraLookAt.x, "Fixed") & "   y: " & Format(m_CameraLookAt.y, "Fixed") & "   z: " & Format(m_CameraLookAt.Z, "Fixed")
    
    ' Show current Camera Zoom value.
    Me.CurrentX = sngX
    Me.Print "Camera Zoom: " & Format(m_CameraZoom, "Fixed")
    
    ' Show current Field Of View value.
    Me.CurrentX = sngX
    Me.Print "Field Of View: " & Format(m_CameraFOV, "Fixed") & "°" & vbNewLine
    
    
    ' Show Demo Status Message.
    If InDemo = True Then
        Me.CurrentX = sngX
        Me.ForeColor = RGB(255, 128, 128)
        If DemoCounter < 150 Then
            Me.Print "Demo Status: The Camera and LookAt vector are both moving."
        ElseIf DemoCounter < 250 Then
            Me.Print "Demo Status: Moving Camera back to start position."
        ElseIf DemoCounter < 350 Then
            Me.Print "Demo Status: Adjusting the Zoom value back to 1.0"
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    ' Set some basic form properties.
    ' ===============================
    Me.AutoRedraw = True
    Me.BackColor = RGB(0, 0, 0)
    
    
    ' Create our test data.
    ' =====================
    Call CreateTestData
    
    
    ' Position the Camera away from the center of the Dots.
    ' =====================================================
    '   * Positive X points to the Right
    '   * Positive Z points *into* the monitor - away from you.
    '   * Positive Y goes Up
    m_Camera.x = -200
    m_Camera.y = 1500
    m_Camera.Z = -200
    m_Camera.W = 1
    
    
    ' Reset the Camera's LookAt point.
    ' ================================
    m_CameraLookAt.x = 100
    m_CameraLookAt.y = 0
    m_CameraLookAt.Z = 0
    m_CameraLookAt.W = 1
    
    
    ' Reset the Camera's Zoom value.
    ' ==============================
    m_CameraZoom = 0.5
    m_CameraFOV = CalculateFOV(m_CameraZoom)
    
    ' Show and Resize the Form
    ' ========================
    Me.Show
    
    
    ' Reset the clipping distances.
    m_ClippingDistanceFar = 500
    m_ClippingDistanceNear = 0
    
    
    ' Hide Mouse (by moving it to the far right)
    ' This method causes less problems than actually hiding the mouse!
    ' ================================================================
    Call SetCursorPos(Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)

    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' This routine slows down the program when the mouse is moved.
    ' ============================================================
    
    On Error GoTo errTrap
    
    Dim lngIndex As Long
    Dim PixelX As Single
    Dim PixelY As Single
    Dim sngAutoTolerance As Single
    
    If Me.TimerMain.Enabled = False Then Exit Sub
    
    Me.DrawWidth = 4
    Me.ForeColor = RGB(255, 255, 0)
    
    sngAutoTolerance = m_CameraZoom / 50
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Temp) To UBound(m_Temp)
        
        ' Only consider dots in front of us.
        If m_Temp(lngIndex).Z > 0 Then
            
            PixelX = m_Temp(lngIndex).x
            PixelY = -m_Temp(lngIndex).y
            
            ' Is the mouse close to the X coordinate?
            If Abs(PixelX - x) < sngAutoTolerance Then

                ' Is the mouse close to the Y coordinate?
                If Abs(PixelY - y) < sngAutoTolerance Then

                    ' Plot the pixel
                    Me.PSet (PixelX, PixelY)
                    Me.Print "x:" & Format(m_Dots(lngIndex).x, "Fixed") & " y:" & Format(m_Dots(lngIndex).y, "Fixed") & " z:" & Format(m_Dots(lngIndex).Z, "Fixed")

                End If
            End If
        End If
     Next lngIndex
     
    Exit Sub
errTrap:
    
End Sub

Private Sub Form_Resize()

    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier.
    
    Me.ScaleWidth = 2 * (1 / m_CameraZoom)
    Me.ScaleLeft = -(1 / m_CameraZoom)
    
    Me.ScaleHeight = Me.ScaleWidth
    Me.ScaleTop = Me.ScaleLeft
    
End Sub
Public Sub CalculateNewDotPositions()

    On Error GoTo errTrap
    
    Dim lngIndex As Long
    
    ReDim m_Temp(UBound(m_Dots))
    
    Dim matView As mdrMATRIX4
    Dim vectVPN As mdrVector4 ' View Plane Normal (VPN) - The direction that the Virtual Camera is pointing.
    Dim vectVUP As mdrVector4 ' View UP direction (VUP) - Which way is up? This is used for tilting (or not tilting) the Camera.
    Dim vectVRP As mdrVector4 ' View Reference Point (VRP) - The World Position of the Virtual Camera.
    
    
    ' Subtract the Camera's world position from the 'LookAt' point to give us the View Plane Normal (VPN).
    vectVPN = VectorSubtract(m_CameraLookAt, m_Camera)
    
    With vectVUP
        .x = 0
        .y = 1
        .Z = 0
        .W = 1
    End With
    
    With vectVRP            ' I'm kind of duplicating things here a bit - don't worry about it.
        .x = m_Camera.x
        .y = m_Camera.y
        .Z = m_Camera.Z
        .W = 1
    End With
    
    ' Calculate the View Orientation Matrix.
    matView = MatrixViewOrientation(vectVPN, vectVUP, vectVRP)
    
    
    ' Loop through from the "Lower Boundry" of the Array, to the "Upper Boundry" of the Array.
    For lngIndex = LBound(m_Dots) To UBound(m_Dots)
    
        ' Apply the 'View Matrix' to the dots.
        m_Temp(lngIndex) = MatrixMultiplyVector(matView, m_Dots(lngIndex))
        
        
        ' ========================================================
        ' Transform the 3D vector down to a 2D vector by simply
        ' dividing the X and Y coordinates by their Z counterpart.
        ' ========================================================
        If m_Temp(lngIndex).Z <> 0 Then
            m_Temp(lngIndex).x = m_Temp(lngIndex).x / m_Temp(lngIndex).Z
            m_Temp(lngIndex).y = m_Temp(lngIndex).y / m_Temp(lngIndex).Z
        End If
        
    Next lngIndex
    
    Exit Sub
errTrap:
    
End Sub


Public Sub UpdateCameraParameters()

    ' ===================================================================
    ' This routine looks at the keyboard, and adjusts the camera position
    ' and Zoom values depending on which keys are held down.
    ' ===================================================================
    
    Dim lngKeyState As Long
    Dim sngCameraStep As Single
    
    sngCameraStep = 1 ' << Adjust this to move the camera faster or slower (any value not zero)
    
    lngKeyState = GetKeyState(vbKeyControl)
    If (lngKeyState And &H8000) Then
    
        ' ======================================
        ' Move Camera's LookAt Point Left/Right.
        ' ======================================
        lngKeyState = GetKeyState(vbKeyLeft)
        If (lngKeyState And &H8000) Then m_CameraLookAt.x = m_CameraLookAt.x - sngCameraStep
        lngKeyState = GetKeyState(vbKeyRight)
        If (lngKeyState And &H8000) Then m_CameraLookAt.x = m_CameraLookAt.x + sngCameraStep
    
    Else
    
        ' =======================
        ' Move Camera Left/Right.
        ' =======================
        lngKeyState = GetKeyState(vbKeyLeft)
        If (lngKeyState And &H8000) Then m_Camera.x = m_Camera.x - sngCameraStep
        lngKeyState = GetKeyState(vbKeyRight)
        If (lngKeyState And &H8000) Then m_Camera.x = m_Camera.x + sngCameraStep
    
        lngKeyState = GetKeyState(vbKeyShift)
        If (lngKeyState And &H8000) Then
            
            ' ======================================================================
            ' Shift Key is down, the user must want to move closer, or further away.
            ' ======================================================================
            lngKeyState = GetKeyState(vbKeyUp)
            If (lngKeyState And &H8000) Then m_Camera.Z = m_Camera.Z + sngCameraStep
            lngKeyState = GetKeyState(vbKeyDown)
            If (lngKeyState And &H8000) Then m_Camera.Z = m_Camera.Z - sngCameraStep
        
        Else
        
            ' =============================================
            ' Shift Key is *not* down. Move camera up/down.
            ' =============================================
            lngKeyState = GetKeyState(vbKeyUp)
            If (lngKeyState And &H8000) Then m_Camera.y = m_Camera.y + sngCameraStep
            lngKeyState = GetKeyState(vbKeyDown)
            If (lngKeyState And &H8000) Then m_Camera.y = m_Camera.y - sngCameraStep
            
        End If ' Is Shift Key held down?
    End If ' Is Control Key held down?
    
    ' ==============================================================================================
    ' Modify the following:
    '   * Field Of View (FOV)
    '   * Camera's Zoom value
    '
    '   Note: These two values are pretty much the same thing, it depends on how you think about it.
    '         You could also think of this as the "Perspective Distortion" as well.
    '
    ' All of this is achieved simply by adjusting the height/width of the window.
    ' It might sound simple, but in reality this is pretty much what the complex 3D engines do.
    ' ==============================================================================================
    lngKeyState = GetKeyState(vbKeyPageUp)
    If (lngKeyState And &H8000) Then
        If m_CameraZoom > 0.05 Then
            m_CameraZoom = m_CameraZoom - 0.05
            m_CameraFOV = CalculateFOV(m_CameraZoom)
            Call Form_Resize                                                '   Redefine the Height/Width of our drawing window.
        End If
    End If
    lngKeyState = GetKeyState(vbKeyPageDown)
    If (lngKeyState And &H8000) Then
        m_CameraZoom = m_CameraZoom + 0.05
        m_CameraFOV = CalculateFOV(m_CameraZoom)
        Call Form_Resize                                                '   Redefine the Height/Width of our drawing window.
    End If
    
    
    ' ====================================
    ' Reset Camera to a starting position.
    ' ====================================
    lngKeyState = GetKeyState(vbKeySpace)
    If (lngKeyState And &H8000) Then
        ' Reset Camera
        m_Camera.x = 0
        m_Camera.y = 3
        m_Camera.Z = -15

        ' Reset the Camera's LookAt point.
        ' ================================
        m_CameraLookAt.x = 0
        m_CameraLookAt.y = 0
        m_CameraLookAt.Z = 0
        
        ' Reset Zoom/FOV
        m_CameraZoom = 1
        m_CameraFOV = CalculateFOV(m_CameraZoom)
        Call Form_Resize
        
    End If
    
    
    ' ========================================
    ' Check for ESC / Quit / Exit Application.
    ' ========================================
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then
        ' Quit Application
        Me.TimerMain.Enabled = False
        Unload Me
    End If
    
    
End Sub

Private Sub TimerMain_Timer()

    ' Apply Virtual Camera to original 'Dots'
    Call CalculateNewDotPositions
    
    ' Draw Stuff
    ' ==========
    Me.Cls
    Call DrawCrossHairs
    Call DrawDots
    Call DrawParameters(False, 0)
    
    ' Process keyboard commands
    Call UpdateCameraParameters
    
End Sub

Private Sub TimerZoomIn_Timer()

    Static lngCounter As Long
    lngCounter = lngCounter + 1
    
    ' Apply Virtual Camera to original 'Dots'
    Call CalculateNewDotPositions
    
    
    ' Draw Stuff
    ' ==========
    Me.Cls
    Call DrawDots
    Call DrawParameters(True, lngCounter)
    
    
    ' Animate our Virtual Camera for a quick introduction.
    ' ====================================================
    Dim vectTemp As mdrVector4, vectTemp2 As mdrVector4
    If lngCounter < 150 Then        ' << Zoom In
        ' Move the 'LookAt' point.
        ' ========================
        vectTemp.x = 0
        vectTemp.y = 0
        vectTemp.Z = 0
        vectTemp = VectorSubtract(vectTemp, m_CameraLookAt)
        m_CameraLookAt.x = m_CameraLookAt.x + (vectTemp.x * 0.06)
        m_CameraLookAt.y = m_CameraLookAt.y + (vectTemp.y * 0.1)
        m_CameraLookAt.Z = m_CameraLookAt.Z + (vectTemp.Z * 0.1)
        
        ' Move the 'Virtual Camera'.
        ' ==========================
        vectTemp2.x = 1.5
        vectTemp2.y = -1.75
        vectTemp2.Z = -1.5
        vectTemp = VectorSubtract(m_CameraLookAt, m_Camera)
        vectTemp = VectorAddition(vectTemp, vectTemp2)
        m_Camera.x = m_Camera.x + (vectTemp.x * 0.1)
        m_Camera.y = m_Camera.y + (vectTemp.y * 0.1)
        m_Camera.Z = m_Camera.Z + (vectTemp.Z * 0.1)
            
    ElseIf lngCounter < 250 Then    ' << Move the Camera back out.
            
        ' Move the 'Virtual Camera'.
        ' ==========================
        vectTemp.x = 0
        vectTemp.y = 3
        vectTemp.Z = -15
        vectTemp = VectorSubtract(vectTemp, m_Camera)
        m_Camera.x = m_Camera.x + (vectTemp.x * 0.03)
        m_Camera.y = m_Camera.y + (vectTemp.y * 0.03)
        m_Camera.Z = m_Camera.Z + (vectTemp.Z * 0.035)
        
    ElseIf lngCounter < 350 Then    ' Restore the Zoom value.
    
        m_CameraLookAt.x = 0
        m_CameraLookAt.y = 0
        m_CameraLookAt.Z = 0
        
        m_Camera.x = 0
        m_Camera.y = 3
        m_Camera.Z = -15
        
        m_CameraZoom = m_CameraZoom - ((m_CameraZoom - 1) / 15)
        m_CameraFOV = CalculateFOV(m_CameraZoom)
        Call Form_Resize
        
    Else
    
        ' Reset LookAt point.
        m_CameraLookAt.x = 0
        m_CameraLookAt.y = 0
        m_CameraLookAt.Z = 0
        
        ' Reset Camera
        m_Camera.x = 0
        m_Camera.y = 3
        m_Camera.Z = -15
        
        ' Reset Zoom Value
        m_CameraZoom = 1
        m_CameraFOV = CalculateFOV(m_CameraZoom)
        Call Form_Resize
        
        ' Reset the clipping distances.
        m_ClippingDistanceFar = 300
        m_ClippingDistanceNear = 0
    
        ' Switch Timers
        Me.TimerZoomIn.Enabled = False
        Me.TimerMain.Enabled = True
    End If
    
    
    ' ======================
    ' Check for ESC of Demo.
    ' ======================
    Dim lngKeyState As Long
    lngKeyState = GetKeyState(vbKeyEscape)
    If (lngKeyState And &H8000) Then lngCounter = 451
    
End Sub

