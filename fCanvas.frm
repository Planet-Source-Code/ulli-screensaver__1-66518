VERSION 5.00
Begin VB.Form fCanvas 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   480
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Timer tmrTick 
      Interval        =   10
      Left            =   30
      Top             =   0
   End
End
Attribute VB_Name = "fCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Point3D
    X               As Double
    Y               As Double
    Z               As Double
    Color           As Long
End Type

Private Type OriginToProjection
    Origin          As Point3D
    Projection      As Point3D
End Type

Private Type Obj3D
    A               As OriginToProjection
    B               As OriginToProjection
    C               As OriginToProjection
    D               As OriginToProjection
    AB              As Point3D
    BC              As Point3D
    CD              As Point3D
    PtCount         As Long
    Pt()            As OriginToProjection
End Type

Private i           As Long
Private px          As Single
Private py          As Single

Private AngleX      As Double
Private AngleY      As Double
Private AngleZ      As Double

Private MyOBJ       As Obj3D

Private Pt          As OriginToProjection

Private Rotation    As Point3D

Private Pi          As Double
Private TwoPi       As Double
Private HalfPi      As Double
Private BSize       As Double

Private Sub Delta(Delta As Point3D, PointS As Point3D, PointE As Point3D)

    With Delta
        .X = PointE.X - PointS.X
        .Y = PointE.Y - PointS.Y
        .Z = PointE.Z - PointS.Z
    End With 'DELTA

End Sub

Private Sub Deltas_3Sides(O3D As Obj3D)

    With O3D
        Delta .AB, .A.Projection, .B.Projection
        Delta .BC, .B.Projection, .C.Projection
        Delta .CD, .C.Projection, .D.Projection
    End With 'O3D

End Sub

Private Sub Form_Click()

    If hThumb = 0 Then 'running as screensaver
        Unload Me
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Form_Click

End Sub

Private Sub Form_Load()

    ScaleMode = vbPixels
    BackColor = vbBlack
    If hThumb Then 'running in thumb window
        DrawWidth = 1
      Else 'HTHUMB = FALSE/0
        DrawWidth = 5
        WindowState = vbMaximized
    End If
    Pi = 4 * Atn(1)
    TwoPi = Pi + Pi
    HalfPi = Pi / 2
    DoEvents
    BSize = Height / 1.8
    With MyOBJ
        .A.Origin.X = -BSize
        .A.Origin.Y = -BSize
        .A.Origin.Z = -BSize
        .B.Origin.X = BSize
        .B.Origin.Y = -BSize
        .B.Origin.Z = -BSize
        .C.Origin.X = BSize
        .C.Origin.Y = -BSize
        .C.Origin.Z = BSize
        .D.Origin.X = BSize
        .D.Origin.Y = BSize
        .D.Origin.Z = BSize
        .A.Projection = .A.Origin
        .B.Projection = .B.Origin
        .C.Projection = .C.Origin
        .D.Projection = .D.Origin
    End With 'MYOBJ

    MyOBJ.PtCount = 1024
    ReDim MyOBJ.Pt(1 To MyOBJ.PtCount)

    For i = 1 To MyOBJ.PtCount
        With MyOBJ.Pt(i).Origin
            AngleZ = Rnd * HalfPi
            .X = Cos(AngleZ)
            .Y = Sin(AngleZ)

            AngleX = Rnd * Pi
            .Z = .Y * Sin(AngleX)
            .Y = .Y * Cos(AngleX)

            AngleY = Rnd * TwoPi
            Pt.Origin.Z = .Z * Cos(AngleY) + .X * Sin(AngleY)
            .X = .X * Cos(AngleY) - .Z * Sin(AngleY)
            .Z = Pt.Origin.Z

            'based on known scalar properties of above points,
            'scale and translate them into (0,0,0) - (1,1,1)
            .Z = .Z * 0.5 + 0.5
            .Y = .Y * 0.5 + 0.5
            .X = .X * 0.5 + 0.5
            MyOBJ.Pt(i).Projection.Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        End With 'MYOBJ.PT(I).ORIGIN

    Next i

    Rotation.X = 0.007
    Rotation.Y = 0.011
    Rotation.Z = 0.013
    tmrTick.Interval = 20
    tmrTick.Enabled = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If px <> 0 And py <> 0 And (Abs(X - px) > 4 Or Abs(Y - py) > 4) Then
        Form_Click
      Else 'NOT PX...
        px = X
        py = Y
    End If

End Sub

Private Sub Project(O3D As Obj3D, ByVal TransX As Double, ByVal TransY As Double, Optional ByVal TransZ As Double = 2, Optional ByVal Scalar As Double)

  Dim LocalPointA As Point3D

    Deltas_3Sides O3D
    With O3D
        LocalPointA = .A.Projection
        LocalPointA.Z = LocalPointA.Z + TransZ
        For i = 1 To .PtCount
            With .Pt(i).Projection
                PSet (.X, .Y), BackColor 'erase previous pt
                ProjectPoint O3D.Pt(i), LocalPointA, O3D.AB, O3D.BC, O3D.CD
                .Z = Scalar / .Z
                .X = .X * .Z + TransX
                .Y = .Y * .Z + TransY
                PSet (.X, .Y), .Color 'draw new
            End With '.PT(I).PROJECTION
        Next i
    End With 'O3D

End Sub

Private Sub ProjectPoint(Prj3D As OriginToProjection, StartA As Point3D, DeltaAB As Point3D, DeltaBC As Point3D, DeltaCD As Point3D)

    With Prj3D
        .Projection.X = StartA.X + .Origin.X * DeltaAB.X + .Origin.Z * DeltaBC.X + .Origin.Y * DeltaCD.X
        .Projection.Y = StartA.Y + .Origin.X * DeltaAB.Y + .Origin.Z * DeltaBC.Y + .Origin.Y * DeltaCD.Y
        .Projection.Z = StartA.Z + .Origin.X * DeltaAB.Z + .Origin.Z * DeltaBC.Z + .Origin.Y * DeltaCD.Z
    End With 'PRJ3D

End Sub

Private Sub Rotate(O3D As Obj3D, Optional ByVal AngleX As Double, Optional ByVal AngleY As Double, Optional ByVal AngleZ As Double)

    With O3D
        RotatePoint .A.Projection, AngleX, AngleY, AngleZ, .A.Origin
        RotatePoint .B.Projection, AngleX, AngleY, AngleZ, .B.Origin
        RotatePoint .C.Projection, AngleX, AngleY, AngleZ, .C.Origin
        RotatePoint .D.Projection, AngleX, AngleY, AngleZ, .D.Origin
    End With 'O3D

End Sub

Private Sub RotatePoint(P3D As Point3D, AngleX As Double, AngleY As Double, AngleZ As Double, P3D_SRC As Point3D)

    With P3D
        .X = P3D_SRC.X * Cos(AngleY) - P3D_SRC.Z * Sin(AngleY)
        .Z = P3D_SRC.Z * Cos(AngleY) + P3D_SRC.X * Sin(AngleY)

        Pt.Projection.X = .X * Cos(AngleZ) - P3D_SRC.Y * Sin(AngleZ)
        .Y = P3D_SRC.Y * Cos(AngleZ) + .X * Sin(AngleZ)
        .X = Pt.Projection.X

        Pt.Projection.Z = .Z * Cos(AngleX) - .Y * Sin(AngleX)
        .Y = .Y * Cos(AngleX) + .Z * Sin(AngleX)
        .Z = Pt.Projection.Z
    End With 'P3D

End Sub

Private Sub tmrtick_Timer()

    Rotate MyOBJ, AngleX, AngleY, AngleZ
    Project MyOBJ, ScaleWidth / 2, ScaleHeight / 2, 350, ScaleHeight / 2
    With Rotation
        AngleX = AngleX + .X
        AngleY = AngleY + .Y
        AngleZ = AngleZ + .Z
    End With 'ROTATION
    If AngleX > TwoPi Then
        AngleX = AngleX - TwoPi
    End If
    If AngleY > TwoPi Then
        AngleY = AngleY - TwoPi
    End If
    If AngleZ > TwoPi Then
        AngleZ = AngleZ - TwoPi
    End If

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Sep-09 14:57)  Decl: 44  Code: 198  Total: 242 Lines
':) CommentOnly: 2 (0,8%)  Commented: 15 (6,2%)  Empty: 52 (21,5%)  Max Logic Depth: 4
