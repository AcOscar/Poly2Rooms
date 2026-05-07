Option Explicit
Type raum
     nummer As String ' Raumnummer
     abt As String ' Abteilung
     ref As String ' Referat
     sach As String ' Sachbearbeiter
End Type

Sub alleareas()
On Error Resume Next
Dim ent, ent1 As AcadEntity
Dim lineObj As AcadLine
Dim intPoints As Variant
Dim rayObj As AcadRay
Dim Point1(0 To 2) As Double
Dim Point2(0 To 2) As Double
Dim mtext As AcadMText

Dim area As AecArea
Dim Zaehler1 As Long, Zaehler2 As Long, Zaehler3 As Long

Dim profil As AecProfile

For Each ent1 In ThisDrawing.ModelSpace
    If TypeOf ent1 Is AecArea Then
        For Each ent In ThisDrawing.ModelSpace
            If TypeOf ent Is AcadMText Then
                Set mtext = ent
                Point1(0) = mtext.InsertionPoint(0): Point1(1) = mtext.InsertionPoint(1): Point1(2) = mtext.InsertionPoint(2)
                Point2(0) = mtext.InsertionPoint(0): Point2(1) = mtext.InsertionPoint(1) + 1: Point2(2) = mtext.InsertionPoint(2)
                Set rayObj = ThisDrawing.ModelSpace.AddRay(Point1, Point2)
                intPoints = rayObj.IntersectWith(ent1, acExtendNone)
                If UBound(intPoints) Mod 2 = 0 Then ' wenn ungerade ist er inerhalb
                    Dim sucher As String
                    Dim oraum As raum
                    
                    oraum.nummer = " "
                    oraum.abt = " "
                    oraum.sach = " "
                    oraum.ref = " "
                    
                    Dim suchpos As Long
                    
                    sucher = mtext.TextString
                    
                    'foramtierung entferenen
                    If Left(sucher, 1) = "{" Then
                        sucher = Mid(sucher, 2, Len(sucher) - 3)
                    End If
                    If Left(sucher, 2) = "\H" Then
                        suchpos = InStr(sucher, ";")
                        sucher = Mid(sucher, suchpos + 1, Len(sucher) - 2)
                    End If
                    If Left(sucher, 7) = "\pt2.1;" Then
                        sucher = Mid(sucher, 8, Len(sucher) - 2)
                    End If

                    'Debug.Print sucher
                    
                    If Len(sucher) > 1 Then '
                        suchpos = InStr(sucher, "\P")
                        If suchpos = 0 Then
                            oraum.nummer = sucher
                            sucher = ""
                        Else
                            oraum.nummer = Left(sucher, suchpos - 1)
                            sucher = Mid(sucher, suchpos + 2)
                        End If
                    End If
                            
                    If Len(sucher) > 1 Then
                        suchpos = InStr(sucher, "\P")
                        If suchpos = 0 Then
                            oraum.abt = sucher
                            sucher = ""
                        Else
                            oraum.abt = Left(sucher, suchpos - 1)
                            sucher = Mid(sucher, suchpos + 2)
                        End If
                    End If
                    
                    If Len(sucher) > 1 Then
                        suchpos = InStr(sucher, "\P")
                        If suchpos = 0 Then
                            oraum.ref = sucher
                            sucher = ""
                        Else
                            oraum.ref = Left(sucher, suchpos - 1)
                            sucher = Mid(sucher, suchpos + 2)
                        End If
                    End If
                    
                    If Len(sucher) > 1 Then
                            oraum.sach = sucher
                            sucher = ""
                    End If
                    'Debug.Print oraum.nummer & "  " & oraum.abt & "  " & oraum.ref & "  " & oraum.sach
                    
                    Set area = ent1 'Set poly = ent1
                    Dim SchedApp As AecScheduleApplication
                    Dim propSets As AecSchedulePropertySets
                    Dim propSet As AecSchedulePropertySet
                    Dim props As AecScheduleProperties
                    Set SchedApp = New AecScheduleApplication
                    Set propSets = SchedApp.PropertySets(ent1)
                    Set propSet = propSets.Item(0)
                    Set props = propSet.Properties
                    Dim Data() As Variant
                    ReDim Data(props.Count, 1)
                    For Zaehler2 = 0 To props.Count - 1
                        On Error Resume Next
                        Data(Zaehler2, 1) = Null
                        Data(Zaehler2, 1) = props(Zaehler2).Value
                        Data(Zaehler2, 0) = props(Zaehler2).Name
                        'Debug.Print props(Zaehler2).Name, props(Zaehler2).Value
                        
                        If props(Zaehler2).Name = "Nummer" Then
                            props(Zaehler2).Value = oraum.nummer
                        End If
                        
                        If props(Zaehler2).Name = "Abteilung" Then
                            props(Zaehler2).Value = oraum.abt
                        End If
                        
                        If props(Zaehler2).Name = "Referat" Then
                            props(Zaehler2).Value = oraum.ref
                        End If
                        
                        If props(Zaehler2).Name = "Nutzung" Then
                            props(Zaehler2).Value = oraum.sach
                        End If
                        
                        
                    Next Zaehler2
                    Zaehler3 = Zaehler3 + 1
                End If
                rayObj.Delete
            End If
        Next
    End If
Next ent1

End Sub

Sub Ausl()
  Dim AecDoc As AecDocument
  Dim Zaehler1 As Long, Zaehler2 As Long, Zaehler3 As Long
  Dim aecObj As AecGeo
  Set AecDoc = AecArchBaseApplication.ActiveDocument
  For Zaehler1 = 0 To AecDoc.ModelSpace.Count - 1
      If TypeName(AecDoc.ModelSpace.Item(Zaehler1)) = "IAecArea" Then
          Set aecObj = ThisDrawing.ModelSpace(Zaehler1)
          Dim SchedApp As AecScheduleApplication
          Dim propSets As AecSchedulePropertySets
          Dim propSet As AecSchedulePropertySet
          Dim props As AecScheduleProperties
          Set SchedApp = New AecScheduleApplication
          Set propSets = SchedApp.PropertySets(aecObj)
          Set propSet = propSets.Item(0)
          Set props = propSet.Properties
          Dim Data() As Variant
          ReDim Data(props.Count, 1)
          For Zaehler2 = 0 To props.Count - 1
              On Error Resume Next
              Data(Zaehler2, 1) = Null
              Data(Zaehler2, 1) = props(Zaehler2).Value
              Data(Zaehler2, 0) = props(Zaehler2).Name
              Debug.Print props(Zaehler2).Name, props(Zaehler2).Value
         
          Next Zaehler2
          Zaehler3 = Zaehler3 + 1
      End If
  Next Zaehler1
  Set AecDoc = Nothing
  If Zaehler3 = 0 Then
      MsgBox "Es sind keine Flächen in dieser Zeichnung definiert." & vbCrLf & _
      "Die Funktion wird abgebrochen", vbOKOnly, "Keine definierten Flächen"
      Exit Sub
  End If
  Set SchedApp = Nothing
  Set propSets = Nothing
  Set propSet = Nothing
  Set props = Nothing
End Sub

Sub allepolys()
    On Error Resume Next
    Dim ent, ent1 As AcadEntity
    Dim lineObj As AcadLine
    Dim intPoints As Variant
    Dim rayObj As AcadRay
    Dim Point1(0 To 2) As Double
    Dim Point2(0 To 2) As Double
    Dim mtext As AcadMText
    Dim poly As AcadLWPolyline
    
    For Each ent1 In ThisDrawing.ModelSpace
        If TypeOf ent1 Is AcadLWPolyline Then
        
            For Each ent In ThisDrawing.ModelSpace
                If TypeOf ent Is AcadMText Then
                    Set mtext = ent
                    Point1(0) = mtext.InsertionPoint(0): Point1(1) = mtext.InsertionPoint(1): Point1(2) = mtext.InsertionPoint(2)
                    Point2(0) = mtext.InsertionPoint(0): Point2(1) = mtext.InsertionPoint(1) + 1: Point2(2) = mtext.InsertionPoint(2)
                    Set rayObj = ThisDrawing.ModelSpace.AddRay(Point1, Point2)
                    intPoints = rayObj.IntersectWith(ent1, acExtendNone)
                    If UBound(intPoints) Mod 2 = 0 Then ' wenn ungerade ist er inerhalb
                        'Debug.Print UBound(intPoints)
                        Set poly = ent1
                        Point2(0) = poly.Coordinates(0): Point2(1) = poly.Coordinates(1): Point2(2) = poly.Coordinates(2)
                        Set lineObj = ThisDrawing.ModelSpace.AddLine(Point1, Point2)
                    End If
                    rayObj.Delete
                End If
            Next
        End If
    Next ent1
End Sub

Public Sub test()
  'On Error Resume Next
  Dim doc As AecBaseDocument
  Dim app As New AecBaseApplication
  Dim cRings As AecRings
  Dim ring1 As AecRing
  Dim ring2 As AecRing
  Dim cProfileStyles As AecProfileStyles
  Dim profileStyle As AecProfileStyle
  Dim profile As AecProfile
  Dim ent As AcadEntity
  Dim pts1 As Variant
  Dim point As AcadPoint
  Dim pt As Variant
  Dim poly As AcadLWPolyline
  Dim oldpoly As AcadPolyline
  Dim pts(0 To 3) As Double
  Dim i As Integer
  
  Set cRings = profile.Rings
  ThisDrawing.Utility.GetEntity ent, pt, "Select polyline"
  If TypeOf ent Is AcadLWPolyline Then
    app.Init ThisDrawing.Application
    Set doc = app.ActiveDocument
  
    Set cProfileStyles = doc.ProfileStyles
    Set profileStyle = cProfileStyles.Item("ringTestProfile")
    If profileStyle Is Nothing Then
      Set profileStyle = cProfileStyles.Add("ringTestProfile")
    End If
    Set profile = profileStyle.profile
    Set poly = ent
    Set oldpoly = ConvertPoly(poly)
    Set ring1 = profile.Rings.Add
    pts1 = poly.Coordinates
    ring1.FromPolyline oldpoly
    
    For Each ent In ThisDrawing.ModelSpace
      If TypeOf ent Is AcadPoint Then
        Set point = ent
        
        pts(0) = point.Coordinates(0): pts(1) = point.Coordinates(1)
        pts(2) = point.Coordinates(0): pts(3) = point.Coordinates(1) + 0.000001
        Set ring2 = profile.Rings.Add
        ring2.FromPoints pts
        
        If ring1.Contains(ring2) = True Then
          point.color = acRed
        Else
          point.color = acBlue
        End If
      End If
    Next
  End If
  oldpoly.Delete
  profileStyle.Delete

End Sub





