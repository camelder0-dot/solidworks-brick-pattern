Attribute VB_Name = "BrickPattern"
' ============================================================================
'  Brick Pattern macro for SOLIDWORKS
'
'  Engraves (cuts) a running-bond brick pattern into a flat face, e.g. the
'  top of a wooden board, ready for CNC routing or laser engraving.
'
'  How to use:
'    1. Open a part and click ONE flat face.
'    2. Run the macro and answer the prompts (all sizes in millimetres).
'
'  Carve styles:
'    1 = Carve the MORTAR lines  -> bricks stay raised, grooves between them
'    2 = Carve the BRICKS        -> bricks are recessed, mortar lines stay raised
'
'  The result is normal Cut-Extrude features in the tree, so you can edit
'  the depth later (double-click the feature) or delete them.
' ============================================================================
Option Explicit

' ---- Default settings (change these if you like) ---------------------------
Private Const DEF_BRICK_LENGTH As Double = 40   ' mm
Private Const DEF_BRICK_HEIGHT As Double = 15   ' mm
Private Const DEF_MORTAR_WIDTH As Double = 2    ' mm
Private Const DEF_CARVE_DEPTH As Double = 1     ' mm
Private Const DEF_CARVE_STYLE As Long = 1       ' 1 = carve mortar, 2 = carve bricks
' ----------------------------------------------------------------------------

Private Const MM As Double = 0.001   ' the SOLIDWORKS API works in metres

Sub main()
    Dim swApp As Object, swModel As Object, swSelMgr As Object, swFace As Object

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Open a part first, then select a flat face.", vbExclamation, "Brick Pattern"
        Exit Sub
    End If
    If swModel.GetType <> 1 Then  ' 1 = swDocPART
        MsgBox "This macro works on parts only.", vbExclamation, "Brick Pattern"
        Exit Sub
    End If

    Set swSelMgr = swModel.SelectionManager
    If swSelMgr.GetSelectedObjectCount2(-1) < 1 Then
        MsgBox "Select ONE flat face (e.g. the top of your board), then run the macro again.", _
               vbExclamation, "Brick Pattern"
        Exit Sub
    End If
    If swSelMgr.GetSelectedObjectType3(1, -1) <> 2 Then ' 2 = swSelFACES
        MsgBox "Select ONE flat face (e.g. the top of your board), then run the macro again.", _
               vbExclamation, "Brick Pattern"
        Exit Sub
    End If
    Set swFace = swSelMgr.GetSelectedObject6(1, -1)
    If Not swFace.GetSurface.IsPlane Then
        MsgBox "The selected face must be flat.", vbExclamation, "Brick Pattern"
        Exit Sub
    End If

    ' ---- Ask for settings ----------------------------------------------------
    Dim brickL As Double, brickH As Double, mortar As Double, depth As Double, styleVal As Double
    If Not AskNumber("Brick length (mm):", DEF_BRICK_LENGTH, brickL) Then Exit Sub
    If Not AskNumber("Brick height (mm):", DEF_BRICK_HEIGHT, brickH) Then Exit Sub
    If Not AskNumber("Mortar line width (mm):", DEF_MORTAR_WIDTH, mortar) Then Exit Sub
    If Not AskNumber("Carve depth into the surface (mm):", DEF_CARVE_DEPTH, depth) Then Exit Sub
    Do
        If Not AskNumber("Carve style:" & vbCrLf & vbCrLf & _
                         "1 = carve the MORTAR lines (bricks stay raised)" & vbCrLf & _
                         "2 = carve the BRICKS (mortar lines stay raised)", _
                         DEF_CARVE_STYLE, styleVal) Then Exit Sub
        If styleVal = 1 Or styleVal = 2 Then Exit Do
        MsgBox "Please enter 1 or 2.", vbExclamation, "Brick Pattern"
    Loop

    brickL = brickL * MM: brickH = brickH * MM: mortar = mortar * MM: depth = depth * MM

    Dim ok As Boolean
    swModel.SketchManager.AddToDB = True
    swModel.SketchManager.DisplayWhenAdded = False

    If styleVal = 1 Then
        ok = CarveMortar(swApp, swModel, swFace, brickL, brickH, mortar, depth)
    Else
        ok = CarveBricks(swApp, swModel, swFace, brickL, brickH, mortar, depth)
    End If

    swModel.SketchManager.AddToDB = False
    swModel.SketchManager.DisplayWhenAdded = True
    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False

    If ok Then
        MsgBox "Brick pattern done!" & vbCrLf & _
               "Tip: double-click the 'Brick' features in the tree to change the depth.", _
               vbInformation, "Brick Pattern"
    Else
        MsgBox "Something went wrong creating the cut. Try a smaller depth or " & _
               "bigger bricks, and make sure the face is flat.", vbExclamation, "Brick Pattern"
    End If
End Sub

' ---------------------------------------------------------------------------
' Style 1: cut the mortar grooves. Horizontal and vertical grooves go in two
' separate cuts so their rectangles never overlap inside one sketch.
' ---------------------------------------------------------------------------
Private Function CarveMortar(swApp As Object, swModel As Object, swFace As Object, _
        brickL As Double, brickH As Double, mortar As Double, depth As Double) As Boolean
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim sk1 As Object, sk2 As Object, sm As Object
    Dim pitchX As Double, pitchY As Double, x As Double, y As Double
    Dim y0 As Double, y1 As Double, overlap As Double, row As Long

    Set sm = swModel.SketchManager
    pitchX = brickL + mortar
    pitchY = brickH + mortar

    ' Sketch 1: horizontal grooves between rows
    Set sk1 = StartSketchOnFace(swApp, swModel, swFace, xMin, xMax, yMin, yMax)
    xMin = xMin - brickL: xMax = xMax + brickL   ' run off the edges cleanly
    y = yMin + brickH
    Do While y < yMax
        sm.CreateCornerRectangle xMin, y, 0, xMax, y + mortar, 0
        y = y + pitchY
    Loop
    sm.InsertSketch True

    ' Sketch 2: vertical joints, staggered every other row
    Set sk2 = StartSketchOnFace(swApp, swModel, swFace, xMin, xMax, yMin, yMax)
    xMin = xMin - brickL: xMax = xMax + brickL
    overlap = mortar / 4   ' poke slightly into the horizontal grooves
    row = 0
    y0 = yMin
    Do While y0 < yMax
        y1 = y0 + brickH
        x = xMin + (row Mod 2) * pitchX / 2 + brickL
        Do While x < xMax
            sm.CreateCornerRectangle x, y0 - overlap, 0, x + mortar, y1 + overlap, 0
            x = x + pitchX
        Loop
        y0 = y0 + pitchY
        row = row + 1
    Loop
    sm.InsertSketch True

    CarveMortar = CutSketch(swModel, sk1, depth, "Brick Mortar Rows") And _
                  CutSketch(swModel, sk2, depth, "Brick Mortar Joints")
End Function

' ---------------------------------------------------------------------------
' Style 2: cut each brick as a pocket; the mortar lines are left standing.
' ---------------------------------------------------------------------------
Private Function CarveBricks(swApp As Object, swModel As Object, swFace As Object, _
        brickL As Double, brickH As Double, mortar As Double, depth As Double) As Boolean
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim sk As Object, sm As Object
    Dim pitchX As Double, pitchY As Double, x As Double, y0 As Double, row As Long

    Set sm = swModel.SketchManager
    pitchX = brickL + mortar
    pitchY = brickH + mortar

    Set sk = StartSketchOnFace(swApp, swModel, swFace, xMin, xMax, yMin, yMax)
    row = 0
    y0 = yMin
    Do While y0 < yMax
        x = xMin - pitchX + (row Mod 2) * pitchX / 2
        Do While x < xMax
            sm.CreateCornerRectangle x, y0, 0, x + brickL, y0 + brickH, 0
            x = x + pitchX
        Loop
        y0 = y0 + pitchY
        row = row + 1
    Loop
    sm.InsertSketch True

    CarveBricks = CutSketch(swModel, sk, depth, "Brick Pockets")
End Function

' Open a new sketch on the face and return the face's extents in sketch coords.
Private Function StartSketchOnFace(swApp As Object, swModel As Object, swFace As Object, _
        ByRef xMin As Double, ByRef xMax As Double, ByRef yMin As Double, ByRef yMax As Double) As Object
    Dim sk As Object
    swModel.ClearSelection2 True
    swFace.Select4 False, Nothing
    swModel.SketchManager.InsertSketch True
    Set sk = swModel.SketchManager.ActiveSketch
    GetFaceBoundsInSketch swApp, swFace, sk, xMin, xMax, yMin, yMax
    Set StartSketchOnFace = sk
End Function

' Prompt for a positive number; returns False if the user cancels.
Private Function AskNumber(prompt As String, defaultVal As Double, ByRef result As Double) As Boolean
    Dim s As String
    Do
        s = InputBox(prompt, "Brick Pattern", CStr(defaultVal))
        If s = "" Then AskNumber = False: Exit Function
        If IsNumeric(s) Then
            If CDbl(s) > 0 Then result = CDbl(s): AskNumber = True: Exit Function
        End If
        MsgBox "Please enter a number greater than 0.", vbExclamation, "Brick Pattern"
    Loop
End Function

' Convert the face's 3D bounding box into the 2D coordinates of the sketch.
Private Sub GetFaceBoundsInSketch(swApp As Object, swFace As Object, swSketch As Object, _
        ByRef xMin As Double, ByRef xMax As Double, ByRef yMin As Double, ByRef yMax As Double)
    Dim box As Variant, xform As Object, mathUtil As Object, pt As Object, d As Variant
    Dim i As Integer, j As Integer, k As Integer, p(2) As Double

    box = swFace.GetBox
    Set xform = swSketch.ModelToSketchTransform
    Set mathUtil = swApp.GetMathUtility

    xMin = 1E+30: yMin = 1E+30: xMax = -1E+30: yMax = -1E+30
    For i = 0 To 1
        For j = 0 To 1
            For k = 0 To 1
                p(0) = box(0 + 3 * i): p(1) = box(1 + 3 * j): p(2) = box(2 + 3 * k)
                Set pt = mathUtil.CreatePoint(p)
                Set pt = pt.MultiplyTransform(xform)
                d = pt.ArrayData
                If d(0) < xMin Then xMin = d(0)
                If d(0) > xMax Then xMax = d(0)
                If d(1) < yMin Then yMin = d(1)
                If d(1) > yMax Then yMax = d(1)
            Next k
        Next j
    Next i
End Sub

' Blind cut-extrude of a sketch into the part. Returns True on success.
Private Function CutSketch(swModel As Object, swSketch As Object, depth As Double, featName As String) As Boolean
    Dim skFeat As Object, cutFeat As Object
    Set skFeat = swSketch          ' a Sketch is also a Feature
    swModel.ClearSelection2 True
    skFeat.Select2 False, 0

    Set cutFeat = swModel.FeatureManager.FeatureCut4(True, False, False, 0, 0, depth, depth, _
        False, False, False, False, 0, 0, False, False, False, False, False, _
        True, True, True, True, False, 0, 0, False, False)

    If cutFeat Is Nothing Then
        CutSketch = False
    Else
        cutFeat.Name = featName
        CutSketch = True
    End If
End Function
