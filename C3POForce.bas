


Sub Main()

Call LightForce
Call C3P0IsTheBest

End Sub

Sub LightForce()
  
  Call SetUp
  Call Glass
  Call Screens
  Call XDash
  Call NonXDash
  Call CleanUp
  
End Sub
  
'=================================================================================================================================================================================================
'INITIAL HIGHLIGHTING & START SHEET 4 WITH HFA
'=================================================================================================================================================================================================
Sub SetUp()
  
  Application.ScreenUpdating = False
  
  'Open new sheets
  Worksheets(2).Copy After:=Worksheets(2)
  Worksheets(3).Name = "Validation"
  Worksheets(3).Range("B1").EntireColumn.Insert
  Worksheets(3).Range("B1").Value = "Validation"
  Sheets.Add After:=Worksheets(3)
  Worksheets(4).Name = "Cuts"
  
  
  'Set up for loop
  Dim glassCount As Integer, hFACutIndex As Integer, cutFlatten As Integer, j As Integer, i As Integer
  Dim hFACutLower As Double, hFACutUpper As Double
  
  glassCount = 1
  hFACutIndex = 1

  Worksheets(3).Columns(17).ClearContents
  Worksheets(3).Columns(18).ClearContents
  Worksheets(3).Range("Q1").Value = "HFA Cut Length"
  
  'OUTER LOOP HFA
  For j = 2 To Worksheets(1).Range("C3000").End(xlUp).Row
    'HFA
    'Highlight Glass
    If InStr(Worksheets(1).Range("E" & j), "GT") = 1 _
    Or InStr(Worksheets(1).Range("E" & j), "GA") = 1 Then
      glassCount = glassCount + 1
      Worksheets(1).Range("A" & j).Interior.Color = rgbOrange
      Worksheets(1).Range("B" & j).Interior.Color = rgbOrange
      Worksheets(1).Range("C" & j).Interior.Color = rgbOrange
      Worksheets(1).Range("D" & j).Interior.Color = rgbOrange
      Worksheets(1).Range("E" & j).Interior.Color = rgbOrange
      Worksheets(3).Range("AZ" & glassCount).Value = Worksheets(1).Range("J" & j).Value
      Worksheets(3).Range("BA" & glassCount).Value = Worksheets(1).Range("K" & j).Value
      'Track Position of HFA for Highlighting Later
      Worksheets(3).Range("BG1").Value = "HFA BOM line"
      Worksheets(3).Range("BG" & glassCount).Value = j
    End If
    'HFA
    'Highlight Frame and other parts that are not relivant and Spacer
    If InStr(Worksheets(1).Range("E" & j), "VIG") = 1 _
    Or InStr(Worksheets(1).Range("E" & j), "VP") = 1 _
    Or InStr(Worksheets(1).Range("E" & j), "FR") = 1 _
    Or InStr(Worksheets(1).Range("E" & j), "FP") = 1 _
    Or InStr(Worksheets(1).Range("E" & j), "WNIG") <> 0 _
    Or InStr(Worksheets(1).Range("E" & j), "PAINTING") <> 0 _
    Or (InStr(Worksheets(1).Range("Q" & j), "SPACER") <> 0 And Worksheets(1).Range("H" & j).Value = "LI") Then
      Worksheets(1).Range("A" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("B" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("C" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("D" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("E" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("F" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("Q" & j).Interior.Color = rgbGrey
      If InStr(Worksheets(1).Range("Q" & j), "SPACER") <> 0 Then
        Worksheets(1).Range("J" & j).Interior.Color = rgbGrey
      End If
    End If
    
    'HFA Minor Parts
    If InStr(Worksheets(1).Range("Q" & j), "CLIP") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "BREATHER TUBE") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "BUTYL") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), " DSE ") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "CLILP") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "DESICCANT") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "GASKET") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "TAPE") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "SASH STOP") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "SETTING BLOCK") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "CAULKING") <> 0 _
    Or InStr(Worksheets(1).Range("E" & j), "WS") = 1 _
    Or InStr(Worksheets(1).Range("Q" & j), "CVR") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "COVER") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "STRIKE") <> 0 Then
      Worksheets(1).Range("A" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("B" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("C" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("D" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("E" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("F" & j).Interior.Color = rgbGrey
      Worksheets(1).Range("Q" & j).Interior.Color = rgbGrey
    End If
    
    'Highlight Grid Descriptions
    If InStr(Worksheets(1).Range("Q" & j), "GRD") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "GRID") <> 0 _
    Or InStr(Worksheets(1).Range("Q" & j), "MUNTIN") <> 0 Then
      Worksheets(1).Range("Q" & j).Interior.Color = RGB(102, 102, 204)
    End If
    
    'HFA
    'NEW Add Cut parts to 4th Worksheet "Cuts"
    If (Worksheets(1).Range("A" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).Range("A" & j).Interior.Color <> rgbOrange _
    And Worksheets(1).Range("J" & j).Value > 1 _
    And Worksheets(1).Range("H" & j).Value = "LI" _
    And InStr(Worksheets(1).Range("Q" & j), "SPACER") = 0) _
 _
    Or (Worksheets(1).Range("A" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).Range("A" & j).Interior.Color <> rgbOrange _
    And Worksheets(1).Range("H" & j).Value = "EA" _
    And InStr(Worksheets(1).Range("Q" & j), "PRECUT V") <> 0) Then
        For cutFlatten = Worksheets(1).Range("F" & j).Value To 1 Step -1
          Worksheets(4).Range("A" & hFACutIndex).Value = j
          Worksheets(4).Range("B" & hFACutIndex).Value = Worksheets(1).Range("E" & j).Value
          Worksheets(4).Range("B" & hFACutIndex).Interior.Color = rgbRed
          hFACutLower = Val(Worksheets(1).Range("J" & j).Value) - 0.0625
          hFACutUpper = Val(Worksheets(1).Range("J" & j).Value) + 0.0625
          Worksheets(4).Range("C" & hFACutIndex).Value = hFACutLower
          Worksheets(4).Range("D" & hFACutIndex).Value = Worksheets(1).Range("J" & j).Value
          Worksheets(4).Range("E" & hFACutIndex).Value = hFACutUpper
          Worksheets(4).Range("F" & hFACutIndex).Value = Worksheets(1).Range("F" & j).Value
          hFACutIndex = hFACutIndex + 1
        Next cutFlatten
    End If
     
'------------------------------------------------------------------------------------
    
    'INNER LOOP FOR ORACLE PARTS
    For i = 2 To Worksheets(3).Range("C3000").End(xlUp).Row
      'Oracle
      'Highlight All UOM EA And LI
      Worksheets(3).Range("N" & i).Interior.Color = rgbGrey
      'Oracle
      'Grey out any lines that will not be matched like Sublines, IGU, PANEL, LABEL, FRAME, and GLASS
      'Or (InStr(Worksheets(3).Range("C" & i), "*") <> 0 And InStr(Worksheets(3).Range("C" & i), "X-") = 0)
      If Worksheets(3).Range("N" & i).Value = "LI" _
      Or InStr(1, Worksheets(3).Range("C" & i), "PANEL") _
      Or InStr(1, Worksheets(3).Range("C" & i), "LABEL") _
      Or InStr(1, Worksheets(3).Range("C" & i), "FRAME") _
      Or InStr(1, Worksheets(3).Range("C" & i), "X-GT") _
      Or InStr(1, Worksheets(3).Range("C" & i), "X-GA") _
      Or InStr(1, Worksheets(3).Range("C" & i), "GA") _
      Or InStr(1, Worksheets(3).Range("C" & i), "GT") Then
        Worksheets(3).Range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("P" & i).Interior.Color = rgbGrey
      End If
      
      'Oracle
      'Grey out Precuts since they are 'EA'
      If InStr(Worksheets(3).Range("D" & i), "PRECUT V") <> 0 Then
        Worksheets(3).Range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("P" & i).Interior.Color = rgbGrey
      End If
      
      'Oracle Minor Parts and Spacer LI
      If InStr(Worksheets(3).Range("D" & i), "CLIP") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "BREATHER TUBE") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "WEEP") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "BUTYL") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "CVR") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "Cover") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), " DSE ") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "DESICCANT") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "GASKET") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "TAPE") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "SASH STOP") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "BLOCK") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "CAULKING") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "STRIKE") <> 0 _
      Or InStr(Worksheets(3).Range("D" & i), "ARGON") <> 0 _
      Or (InStr(Worksheets(3).Range("D" & i), "SPACER") <> 0 And Worksheets(3).Range("N" & i).Value = "LI") Then
        Worksheets(3).Range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).Range("P" & i).Interior.Color = rgbGrey
      End If
      'Incriment inner loop
    Next i
  'Incriment outer loop
  Next j
  'End of Loop Block
  
End Sub
  
'=================================================================================================================================================================================================
  'GLASS
'=================================================================================================================================================================================================
Sub Glass()
  
  'ORACLE GLASS LOOP
  'Building up array of AW and AX with Oracle glass sizes
  Worksheets(3).Range("AW1").Value = "Oracle GLASS W"
  Worksheets(3).Range("AX1").Value = "Oracle GLASS H"
  
  Dim c As Integer, widthStopCount As Integer
  Dim glassWidthStr As String, glassHeightStr As String
  Dim widthStop As Boolean
  Dim glassWidth As Double, glassHeight As Double
  
  glassCount = 1
  
  For i = 2 To Worksheets(3).Range("C3000").End(xlUp).Row
    'Parse out Width And Height from IGU line And Print to new Column
    If InStr(1, Worksheets(3).Range("C" & i), "IGU") Then
      glassCount = glassCount + 1
      glassWidthStr = ""
      glassHeightStr = ""
      widthStop = False
      widthStopCount = 0
          
      'Look in Description to get width and height
      For c = 1 To Len(Worksheets(3).Range("D" & i).Value)
        Dim currentChar As String
        currentChar = Mid(Worksheets(3).Range("D" & i).Value, c, 1)
        'Build up width measurement
        If widthStop = False Then
          If currentChar = " " Then
            widthStop = True
            widthStopCount = c
          Else
            glassWidthStr = glassWidthStr + currentChar
          End If
        End If
        'Print results in AW AX AY
        If widthStop = True And c > widthStopCount + 2 Then
          If currentChar = " " Then
            glassWidth = Val(glassWidthStr)
            glassHeight = Val(glassHeightStr)
            Worksheets(3).Range("AW" & glassCount).Value = glassWidth
            Worksheets(3).Range("AX" & glassCount).Value = glassHeight
            Worksheets(3).Range("AY1").Value = "Oracle BOM line"
            Worksheets(3).Range("AY" & glassCount).Value = i
            Exit For
          Else
            glassHeightStr = glassHeightStr + currentChar
          End If
        End If
      Next c
    End If
  Next i
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  
  'HFA GLASS LOOP
  'New Loop to remove duplicates for matching Glass, so that only IGU remains.  Modifying Array in AW and AX while adding AY
  Worksheets(3).Range("AZ1").Value = "HFA GLASS W"
  Worksheets(3).Range("BA1").Value = "HFA GLASS H"
  
  Dim lastrowHFAGlass As Long
  Dim position As Integer
  
  lastrowHFAGlass = Worksheets(3).Range("AZ50").End(xlUp).Row
  For i = 2 To lastrowHFAGlass
    For j = 3 To lastrowHFAGlass
      If i <> j _
      And Worksheets(3).Range("AZ" & i).Value = Worksheets(3).Range("AZ" & j).Value _
      And Worksheets(3).Range("BA" & i).Value = Worksheets(3).Range("BA" & j).Value _
      And Worksheets(3).Range("BB" & j).Value <> "flagged" _
      And IsEmpty(Worksheets(3).Range("AY" & j).Value) Then
        Worksheets(3).Range("BB" & i).Value = "flagged"
        Worksheets(3).Range("AZ" & j).Value = "Duplicate"
        Worksheets(3).Range("BA" & j).Value = "Duplicate"
        'Grey out HFA sheet measurement for Duplicate
        position = Worksheets(3).Range("BG" & j).Value
        Worksheets(1).Range("J" & position).Interior.Color = rgbGrey
        Worksheets(1).Range("K" & position).Interior.Color = rgbGrey
        Exit For
      End If
    Next j
  Next i
  
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  'GLASS COMPARE
  'New Loop to match and highlight comparison between Glass array in AW and AX to Description in Oracle
  Dim lastrowOracleGlass As Long
  Dim hfaWidthCut As String, hfaHeightCut As String
  Dim hFAWidthCutLower As Double, hFAWidthCutUpper As Double, hFAHeightCutLower As Double, hFAHeightCutUpper As Double
  
  lastrowOracleGlass = Worksheets(3).Range("AW50").End(xlUp).Row
   
  For j = 2 To lastrowHFAGlass
    hfaWidthCut = Worksheets(3).Range("AZ" & j).Value
    hFAWidthCutLower = Val(Worksheets(3).Range("AZ" & j).Value) - 0.0625
    hFAWidthCutUpper = Val(Worksheets(3).Range("AZ" & j).Value) + 0.0625
        
    hfaHeightCut = Worksheets(3).Range("BA" & j).Value
    hFAHeightCutLower = Val(Worksheets(3).Range("BA" & j).Value) - 0.0625
    hFAHeightCutUpper = Val(Worksheets(3).Range("BA" & j).Value) + 0.0625
    
    Worksheets(3).Range("BC" & j).Value = hFAWidthCutLower
    Worksheets(3).Range("BD" & j).Value = hFAWidthCutUpper
    Worksheets(3).Range("BE" & j).Value = hFAHeightCutLower
    Worksheets(3).Range("BF" & j).Value = hFAHeightCutUpper
  
    For i = 2 To lastrowOracleGlass
      'See if both W and H match
      If Worksheets(3).Range("AW" & i).Value >= hFAWidthCutLower _
      And Worksheets(3).Range("AW" & i).Value <= hFAWidthCutUpper _
      And Worksheets(3).Range("AX" & i).Value >= hFAHeightCutLower _
      And Worksheets(3).Range("AX" & i).Value <= hFAHeightCutUpper _
      And Worksheets(3).Range("AY" & i).Interior.Color <> rgbGreen Then
        position = Worksheets(3).Range("AY" & i).Value
        Worksheets(3).Range("AY" & i).Interior.Color = rgbGreen
        Worksheets(3).Range("D" & position).Interior.Color = rgbGreen
        Worksheets(3).Range("A" & position).Interior.Color = rgbAqua
        Worksheets(3).Range("B" & position).Interior.Color = rgbAqua
        Worksheets(3).Range("C" & position).Interior.Color = rgbAqua
        Worksheets(3).Range("P" & position).Interior.Color = rgbAqua
        Worksheets(3).Range("B" & position).Value = "Correct Size"
        'Input match for glass in Q if there is a match
        Worksheets(3).Range("Q" & position).Value = hfaWidthCut + " X " + hfaHeightCut
        Worksheets(3).Range("Q" & position).Interior.Color = rgbGreen
        'Highlight HFA Page with correct glass
        position = Worksheets(3).Range("BG" & i).Value
        Worksheets(1).Range("J" & position).Interior.Color = rgbGreen
        Worksheets(1).Range("K" & position).Interior.Color = rgbGreen
        Exit For
      End If
    Next i
  Next j
  
  'Glass Cleanup highlight bad glass
  'HFA
  For j = 2 To Worksheets(1).Range("A3000").End(xlUp).Row
    If Worksheets(1).Range("E" & j).Interior.Color = rgbOrange _
    And Worksheets(1).Range("J" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).Range("J" & j).Interior.Color <> rgbGreen Then
       Worksheets(1).Range("J" & j).Interior.Color = rgbSalmon
       Worksheets(1).Range("K" & j).Interior.Color = rgbSalmon
    End If
  Next j
  
  For j = 2 To Worksheets(3).Range("C3000").End(xlUp).Row
    If InStr(1, Worksheets(3).Range("C" & j), "IGU") <> 0 _
    And Worksheets(3).Range("C" & j).Interior.Color <> rgbAqua Then
       Worksheets(3).Range("A" & j).Interior.Color = rgbSalmon
       Worksheets(3).Range("B" & j).Interior.Color = rgbSalmon
       Worksheets(3).Range("C" & j).Interior.Color = rgbSalmon
       Worksheets(3).Range("D" & j).Interior.Color = rgbSalmon
       Worksheets(3).Range("P" & j).Interior.Color = rgbSalmon
    End If
  Next j
  
End Sub
  
'=================================================================================================================================================================================================
'SCREENS
'=================================================================================================================================================================================================
 Sub Screens()
  
  'HFA
  Dim screenCount As Integer, xMarker As Integer, screenFlatten As Integer
  screenCount = 1
  
  For j = 2 To Worksheets(1).Range("C3000").End(xlUp).Row
    If Worksheets(1).Range("H" & j).Value = "EA" _
    And Worksheets(1).Range("K" & j).Value <> 0 Then
      For screenFlatten = Worksheets(1).Range("F" & j).Value To 1 Step -1
        Worksheets(4).Range("AD" & screenCount).Value = j
        Worksheets(4).Range("AE" & screenCount).Value = Worksheets(1).Range("E" & j).Value
        Worksheets(4).Range("AF" & screenCount).Value = Worksheets(1).Range("J" & j).Value - 0.0625
        Worksheets(4).Range("AG" & screenCount).Value = Worksheets(1).Range("J" & j).Value
        Worksheets(4).Range("AH" & screenCount).Value = Worksheets(1).Range("J" & j).Value + 0.0625
        Worksheets(4).Range("AI" & screenCount).Value = Worksheets(1).Range("K" & j).Value - 0.0625
        Worksheets(4).Range("AJ" & screenCount).Value = Worksheets(1).Range("K" & j).Value
        Worksheets(4).Range("AK" & screenCount).Value = Worksheets(1).Range("K" & j).Value + 0.0625
        screenCount = screenCount + 1
      Next screenFlatten
    End If
  Next j
  
  'Oracle
  screenCount = 1
  For j = 2 To Worksheets(3).Range("C3000").End(xlUp).Row
    If InStr(1, Worksheets(3).Range("D" & j).Value, "Screen,") Then
      Worksheets(4).Range("AP" & screenCount).Value = j
      Worksheets(4).Range("AQ" & screenCount).Value = Worksheets(3).Range("C" & j).Value
      Worksheets(4).Range("AR" & screenCount).Value = Worksheets(3).Range("D" & j).Value
      'Width
      xMarker = InStr(1, Worksheets(4).Range("AR" & screenCount).Value, "x")
      'Hundreds Place
      If IsNumeric(Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker - 9, 1)) Then
        Worksheets(4).Range("AS" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker - 9, 8)
        Worksheets(4).Range("AS" & screenCount).Interior.Color = RGB(102, 102, 204)
      'Tens Place
      ElseIf IsNumeric(Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker - 8, 1)) Then
        Worksheets(4).Range("AS" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker - 8, 7)
        Worksheets(4).Range("AS" & screenCount).Interior.Color = rgbGreen
      'Ones Place
      Else
        Worksheets(4).Range("AS" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker - 7, 6)
        Worksheets(4).Range("AS" & screenCount).Interior.Color = rgbGold
      End If
      'Height
      'Hundreds Place
      If IsNumeric(Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker + 9, 1)) Then
        Worksheets(4).Range("AT" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker + 2, 8)
        Worksheets(4).Range("AT" & screenCount).Interior.Color = RGB(102, 102, 204)
      'Tens Place
      ElseIf IsNumeric(Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker + 8, 1)) Then
        Worksheets(4).Range("AT" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker + 2, 7)
        Worksheets(4).Range("AT" & screenCount).Interior.Color = rgbGreen
      'Ones Place
      Else
        Worksheets(4).Range("AT" & screenCount).Value = Mid(Worksheets(4).Range("AR" & screenCount).Value, xMarker + 2, 6)
        Worksheets(4).Range("AT" & screenCount).Interior.Color = rgbGold
      End If
      screenCount = screenCount + 1
    End If
  Next j
  
  'Perfect Match
  'Compare Oracle to HFA
  screenCount = 1
  If Worksheets(4).Range("AQ1").Value <> "" And Worksheets(4).Range("AE1").Value <> "" Then
    For j = 1 To Worksheets(4).Range("AQ50").End(xlUp).Row
      For i = 1 To Worksheets(4).Range("AE50").End(xlUp).Row
        If InStr(1, Worksheets(4).Range("AQ" & j).Value, Worksheets(4).Range("AE" & i).Value) <> 0 _
        And Worksheets(4).Range("AS" & j).Value >= Worksheets(4).Range("AF" & i).Value _
        And Worksheets(4).Range("AS" & j).Value <= Worksheets(4).Range("AH" & i).Value _
        And Worksheets(4).Range("AT" & j).Value >= Worksheets(4).Range("AI" & i).Value _
        And Worksheets(4).Range("AT" & j).Value <= Worksheets(4).Range("AK" & i).Value _
        And Worksheets(4).Range("AQ" & j).Interior.Color <> rgbGreen _
        And Worksheets(4).Range("AE" & i).Interior.Color <> rgbGreen Then
          'Page 4
          Worksheets(4).Range("AQ" & j).Interior.Color = rgbGreen
          Worksheets(4).Range("AE" & i).Interior.Color = rgbGreen
          'Page 3
          Worksheets(3).Range("A" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).Range("B" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).Range("C" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).Range("D" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).Range("P" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).Range("Q" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).Range("Q" & Worksheets(4).Range("AP" & j).Value).Value = Worksheets(4).Range("AG" & i).Value & " X " & Worksheets(4).Range("AJ" & i).Value
          'Page 1
          Worksheets(1).Range("A" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("B" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("C" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("D" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("E" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("F" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("J" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbGreen
          Worksheets(1).Range("K" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbGreen
        End If
      Next i
    Next j
    
    'Match If Cuts are equal but part is not
    'Compare Oracle to HFA
    screenCount = 1
    For j = 1 To Worksheets(4).Range("AQ50").End(xlUp).Row
      For i = 1 To Worksheets(4).Range("AE50").End(xlUp).Row
        If Worksheets(4).Range("AS" & j).Value >= Worksheets(4).Range("AF" & i).Value _
        And Worksheets(4).Range("AS" & j).Value <= Worksheets(4).Range("AH" & i).Value _
        And Worksheets(4).Range("AT" & j).Value >= Worksheets(4).Range("AI" & i).Value _
        And Worksheets(4).Range("AT" & j).Value <= Worksheets(4).Range("AK" & i).Value _
        And Worksheets(4).Range("AQ" & j).Interior.Color <> rgbGreen _
        And Worksheets(4).Range("AE" & i).Interior.Color <> rgbGreen Then
          'Page 4
          Worksheets(4).Range("AQ" & j).Interior.Color = rgbGreen
          Worksheets(4).Range("AE" & i).Interior.Color = rgbGreen
          'Page 3
          Worksheets(3).Range("A" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).Range("B" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).Range("C" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).Range("D" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).Range("P" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).Range("Q" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).Range("Q" & Worksheets(4).Range("AP" & j).Value).Value = Worksheets(4).Range("AG" & i).Value & " X " & Worksheets(4).Range("AJ" & i).Value
          'Page 1
          Worksheets(1).Range("A" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).Range("B" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).Range("C" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).Range("D" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).Range("E" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).Range("F" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).Range("J" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbGreen
          Worksheets(1).Range("K" & Worksheets(4).Range("AD" & i).Value).Interior.Color = rgbGreen
        End If
      Next i
    Next j
    
    'CleanUp
    'Oracle
    For j = 1 To Worksheets(4).Range("AQ50").End(xlUp).Row
      If Worksheets(4).Range("AQ" & j).Interior.Color <> rgbGreen Then
        Worksheets(3).Range("A" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).Range("B" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).Range("C" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).Range("D" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).Range("P" & Worksheets(4).Range("AP" & j).Value).Interior.Color = rgbSalmon
      End If
    Next j
    'HFA
    For j = 1 To Worksheets(4).Range("AE50").End(xlUp).Row
      If Worksheets(4).Range("AE" & j).Interior.Color <> rgbGreen Then
        Worksheets(1).Range("A" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("B" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("C" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("D" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("E" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("F" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("J" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).Range("K" & Worksheets(4).Range("AD" & j).Value).Interior.Color = rgbSalmon
      End If
    Next j
  End If
  
End Sub
   
'=================================================================================================================================================================================================
'X- PART COMPARE
'=================================================================================================================================================================================================
Sub XDash()

  Dim oracleCutIndex As Integer, oracleFlatten As Integer, y As Integer, z As Integer, b As Integer, startIndex As Integer, strIndexCount As Integer
  Dim match As Boolean
  Dim starSlice As String, oracleCutStr As String, descriptionStr As String, buildingString As String
  Dim hasRun As Boolean
  Dim oracleCut As Double

  oracleCutIndex = 1
     
  'Loop to get all X- parts from Oracle
  For i = 2 To Worksheets(3).Range("C3000").End(xlUp).Row
      If Worksheets(3).Range("A" & i).Interior.Color <> rgbGrey _
      And InStr(Worksheets(3).Range("C" & i), "X-") <> 0 Then
        hasRun = False
        startIndex = 0
        strIndexCount = 0
        buildingString = ""
        For b = 1 To Len(Worksheets(3).Range("D" & i).Value)
          Dim currentDChar As String
          currentDChar = Mid(Worksheets(3).Range("D" & i).Value, b, 1)
          buildingString = buildingString + currentDChar
          If IsNumeric(currentDChar) = True And InStr(buildingString, "Cut to") <> 0 Then
            strIndexCount = strIndexCount + 1
            If hasRun = False Then
              startIndex = b - 1
              hasRun = True
            End If
          End If
        Next b
        
        'Slice X- to get just the part
        starSlice = ""
        For j = (InStr(1, Worksheets(3).Range("C" & i).Value, "X-") + 2) To Len(Worksheets(3).Range("C" & i).Value)
          If Mid(Worksheets(3).Range("C" & i).Value, j, 1) = "*" Then
            Exit For
          End If
          starSlice = starSlice + Mid(Worksheets(3).Range("C" & i).Value, j, 1)
        Next j
        
        descriptionStr = Worksheets(3).Range("D" & i).Value
        oracleCutStr = Mid(descriptionStr, startIndex, strIndexCount + 1)
        oracleCut = Val(oracleCutStr)
    
        For oracleFlatten = Worksheets(3).Range("P" & i).Value To 1 Step -1
          Worksheets(4).Range("J" & oracleCutIndex).Value = i
          Worksheets(4).Range("K" & oracleCutIndex).Value = Worksheets(3).Range("C" & i).Value
          Worksheets(4).Range("L" & oracleCutIndex).Value = oracleCut
          Worksheets(4).Range("M" & oracleCutIndex).Value = Worksheets(3).Range("P" & i).Value
          Worksheets(4).Range("N" & oracleCutIndex).Value = starSlice
          oracleCutIndex = oracleCutIndex + 1
        Next oracleFlatten
      End If
  Next i
  
 'Sort Oracle list A-Z
  Worksheets(4).Range("J1:N100").Sort Key1:=Range("N1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  'PERFECT MATCH - GREEN
  'Oracle to HFA comparison match
  'All HFA column starts out RED
  For y = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    match = False
    For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
      If Worksheets(4).Range("N" & y).Value = Worksheets(4).Range("B" & z).Value _
      And Worksheets(4).Range("B" & z).Interior.Color <> rgbGreen _
      And ((Worksheets(4).Range("L" & y).Value >= Worksheets(4).Range("C" & z).Value And Worksheets(4).Range("L" & y).Value <= Worksheets(4).Range("E" & z).Value) Or Worksheets(4).Range("D" & z).Value = 1) Then
        Worksheets(4).Range("N" & y).Interior.Color = rgbGreen
        Worksheets(4).Range("B" & z).Interior.Color = rgbGreen
        match = True
        Exit For
      End If
    Next z
  Next y
  
 'BAD CUTS - BLUE
 'Loop to compare unmatched Oracle column to HFA. Identify items that match but the lengths do not.
 'This will catch any bad cuts and before addressing quantity
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.Color <> rgbGreen _
    And Worksheets(4).Range("N" & z).Interior.Color <> rgbGold Then
      For y = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
        If Worksheets(4).Range("B" & y).Value = Worksheets(4).Range("N" & z).Value _
        And Worksheets(4).Range("B" & y).Interior.Color = rgbRed Then
          Worksheets(4).Range("N" & z).Interior.Color = RGB(102, 102, 204)
          Worksheets(4).Range("B" & y).Interior.Color = RGB(102, 102, 204)
          Exit For
        End If
      Next y
    End If
  Next z
  
  'BAD QUANTITY - GOLD
  'Check "HFA" Column for ID's that are both green and blank due to unmatched "HFA" lines
  'Partial approved lines due to quantity off
  For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
    If Worksheets(4).Range("B" & z).Interior.Color <> rgbGreen Then
      For y = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
        If Worksheets(4).Range("A" & y).Value = Worksheets(4).Range("A" & z).Value _
        And Worksheets(4).Range("B" & y).Interior.Color = rgbGreen Then
          'Change from blank to gold
          Worksheets(4).Range("B" & z).Interior.Color = rgbGold
          Exit For
        End If
      Next y
    End If
  Next z
  'Check "Oracle" Column for ID's that are both green and blank due to unmatched "Oracle" lines
  'Partial approved lines due to quantity off
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.Color <> rgbGreen Then
      For y = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
        If Worksheets(4).Range("J" & y).Value = Worksheets(4).Range("J" & z).Value _
        And Worksheets(4).Range("N" & y).Interior.Color = rgbGreen Then
          'Change from blank to gold
          Worksheets(4).Range("N" & z).Interior.Color = rgbGold
          Exit For
        End If
      Next y
    End If
  Next z
 
'------------------------------------------------------------------------------------------------
  'ORACLE
  'Hightlight Validation page with Green
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.Color = rgbGreen Then
      Worksheets(3).Range("D" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbGreen
      Worksheets(3).Range("A" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("B" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("C" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("P" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Blue
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.Color = RGB(102, 102, 204) Then
      Worksheets(3).Range("D" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("A" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("B" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("C" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("P" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Gold
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.Color = rgbGold Then
      Worksheets(3).Range("A" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("B" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("C" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("P" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      If Worksheets(3).Range("D" & Worksheets(4).Range("J" & z).Value).Interior.Color <> rgbGreen Then
        Worksheets(3).Range("D" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      End If
    End If
  Next z
  
  'Hightlight Validation page with Blank
  For z = 1 To Worksheets(4).Range("N3000").End(xlUp).Row
    If Worksheets(4).Range("N" & z).Interior.ColorIndex = xlNone Then
      Worksheets(3).Range("A" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("B" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("C" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("D" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("P" & Worksheets(4).Range("J" & z).Value).Interior.Color = rgbSalmon
    
    End If
  Next z
  
'----------------------------------------------------------------------------------------------
  'HFA
  'Hightlight Validation page with Green
  For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
    If Worksheets(4).Range("B" & z).Interior.Color = rgbGreen Then
      Worksheets(1).Range("J" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbGreen
      Worksheets(1).Range("A" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("B" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("C" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("D" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("E" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("F" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Blue
  For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
    If Worksheets(4).Range("B" & z).Interior.Color = RGB(102, 102, 204) Then
      Worksheets(1).Range("J" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("A" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("B" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("C" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("D" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("E" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("F" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Gold
  For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
    If Worksheets(4).Range("B" & z).Interior.Color = rgbGold Then
      Worksheets(1).Range("A" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("B" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("C" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("D" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("E" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("F" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      If Worksheets(1).Range("D" & Worksheets(4).Range("A" & z).Value).Interior.Color <> rgbGreen Then
        Worksheets(1).Range("J" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      End If
    End If
  Next z
  
  'Hightlight Validation page with Red
  For z = 1 To Worksheets(4).Range("B3000").End(xlUp).Row
    If Worksheets(4).Range("B" & z).Interior.Color = rgbRed Then
      Worksheets(1).Range("A" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("B" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("C" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("D" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("E" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("F" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("J" & Worksheets(4).Range("A" & z).Value).Interior.Color = rgbSalmon
    End If
  Next z
  
End Sub
  
'=================================================================================================================================================================================================
'NON X- PARTS
'=================================================================================================================================================================================================
 Sub NonXDash()
  
  Dim finalFlatten As Integer
  finalFlatten = 1

  'Non-LI remaining parts
  'Check for Perfect Match Parts and Quantity
  'Oracle
  For z = 2 To Worksheets(3).Range("A3000").End(xlUp).Row
    If Worksheets(3).Range("C" & z).Interior.ColorIndex = xlNone Then
      'HFA
      For y = 2 To Worksheets(1).Range("A3000").End(xlUp).Row
        If Worksheets(1).Range("E" & y).Interior.ColorIndex = xlNone _
        And InStr(1, Worksheets(3).Range("C" & z).Value, Worksheets(1).Range("E" & y).Value) <> 0 _
        And InStr(1, Worksheets(3).Range("P" & z).Value, Worksheets(1).Range("F" & y).Value) <> 0 Then
        'Or InStr(1, Worksheets(1).Range("F" & y).Value, ".") <> 0) Then
          Worksheets(3).Range("A" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("B" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("C" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("D" & z).Interior.Color = rgbAqua
          'If InStr(1, Worksheets(1).Range("F" & y).Value, ".") Then
            'Worksheets(3).Range("P" & z).Interior.Color = rgbGrey
          'Else
            Worksheets(3).Range("P" & z).Interior.Color = rgbAqua
          'End If
          
          'HFA
          Worksheets(1).Range("A" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("B" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("C" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("D" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("E" & y).Interior.Color = rgbAqua
          'If InStr(1, Worksheets(1).Range("F" & y).Value, ".") Then
            'Worksheets(1).Range("F" & y).Interior.Color = rgbGrey
          'Else
            Worksheets(1).Range("F" & y).Interior.Color = rgbAqua
          'End If
          Exit For
        End If
      Next y
    End If
  Next z
  
  'HFA Flatten Remaining Parts
  'EA
  For y = 2 To Worksheets(1).Range("A3000").End(xlUp).Row
    If Worksheets(1).Range("C" & y).Interior.ColorIndex = xlNone _
    And InStr(1, Worksheets(1).Range("H" & y).Value, "EA") = 1 Then
      For i = Worksheets(1).Range("F" & y).Value To 1 Step -1
        Worksheets(4).Range("R" & finalFlatten).Value = y
        Worksheets(4).Range("S" & finalFlatten).Value = Worksheets(1).Range("E" & y).Value
        Worksheets(4).Range("T" & finalFlatten).Value = Worksheets(1).Range("F" & y).Value
        finalFlatten = finalFlatten + 1
      Next i
    End If
  Next y
  
  'Oracle Flatten Remaining Parts
  'EA
  finalFlatten = 1
  For y = 2 To Worksheets(3).Range("A3000").End(xlUp).Row
    If Worksheets(3).Range("A" & y).Interior.ColorIndex = xlNone _
    And InStr(1, Worksheets(3).Range("N" & y).Value, "EA") = 1 Then
      For i = Worksheets(3).Range("P" & y).Value To 1 Step -1
        Worksheets(4).Range("X" & finalFlatten).Value = y
        Worksheets(4).Range("Y" & finalFlatten).Value = Worksheets(3).Range("C" & y).Value
        Worksheets(4).Range("Z" & finalFlatten).Value = Worksheets(3).Range("P" & y).Value
        finalFlatten = finalFlatten + 1
      Next i
    End If
  Next y
  
  'Sort Oracle list A-Z
  Worksheets(4).Range("X1:Z100").Sort Key1:=Range("Y1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
  'Oracle compare to HFA
  For y = 1 To Worksheets(4).Range("Y3000").End(xlUp).Row
    For z = 1 To Worksheets(4).Range("S3000").End(xlUp).Row
      If InStr(Worksheets(4).Range("Y" & y).Value, Worksheets(4).Range("S" & z).Value) <> 0 _
      And Worksheets(4).Range("Y" & y).Interior.Color <> rgbOrange _
      And Worksheets(4).Range("S" & z).Interior.Color <> rgbOrange Then
        Worksheets(4).Range("Y" & y).Interior.Color = rgbOrange
        Worksheets(4).Range("S" & z).Interior.Color = rgbOrange
        Exit For
      End If
    Next z
  Next y
  
  'Loop through lines on HFA and Oracle side to mark the hits and non-hits
  'Since lines are in order, if a line ID is only half valid, it will simply turn the validation or HFA sheet red
  'Oracle
  For y = 1 To Worksheets(4).Range("Y3000").End(xlUp).Row
    If Worksheets(4).Range("Y" & y).Interior.Color = rgbOrange Then
      Worksheets(3).Range("A" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("B" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("C" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("D" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).Range("P" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbAqua
    ElseIf Worksheets(4).Range("Y" & y).Value <> "" Then
      Worksheets(3).Range("A" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("B" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("C" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("D" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).Range("P" & Worksheets(4).Range("X" & y).Value).Interior.Color = rgbSalmon
    End If
  Next y
  
  For y = 1 To Worksheets(4).Range("S3000").End(xlUp).Row
    If Worksheets(4).Range("S" & y).Interior.Color = rgbOrange Then
      Worksheets(1).Range("A" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("B" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("C" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("D" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("E" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).Range("F" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbAqua
    ElseIf Worksheets(4).Range("S" & y).Value <> "" Then
      Worksheets(1).Range("A" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("B" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("C" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("D" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("E" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).Range("F" & Worksheets(4).Range("R" & y).Value).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Try one final pass through for HFA parts that are Red and Oracle Parts that got Ignored "grey"
  'This will pick up any Decimal Values missed aka lineals
  '1/16th tolerance
  'HFA
  For y = 2 To Worksheets(1).Range("A3000").End(xlUp).Row
    If Worksheets(1).Range("E" & y).Interior.Color = rgbSalmon Then
      'ORACLE
      For z = 2 To Worksheets(3).Range("A3000").End(xlUp).Row
        If Worksheets(3).Range("A" & z).Interior.Color = rgbGrey _
        And InStr(1, Worksheets(3).Range("C" & z).Value, Worksheets(1).Range("E" & y).Value) _
        And Worksheets(3).Range("P" & z).Value >= (Worksheets(1).Range("F" & y).Value - 0.0625) _
        And Worksheets(3).Range("P" & z).Value <= (Worksheets(1).Range("F" & y).Value + 0.0625) Then
          Worksheets(1).Range("A" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("B" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("C" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("D" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("E" & y).Interior.Color = rgbAqua
          Worksheets(1).Range("F" & y).Interior.Color = rgbAqua
          Worksheets(3).Range("A" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("B" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("C" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("D" & z).Interior.Color = rgbAqua
          Worksheets(3).Range("P" & z).Interior.Color = rgbAqua
          Exit For
        End If
      Next z
    End If
  Next y
  
  'Highlight the Unmatched Red in Oracle
  For y = 2 To Worksheets(3).Range("A3000").End(xlUp).Row
    If Worksheets(3).Range("A" & y).Interior.ColorIndex = xlNone Then
      Worksheets(3).Range("A" & y).Interior.Color = rgbSalmon
      Worksheets(3).Range("B" & y).Interior.Color = rgbSalmon
      Worksheets(3).Range("C" & y).Interior.Color = rgbSalmon
      Worksheets(3).Range("D" & y).Interior.Color = rgbSalmon
      Worksheets(3).Range("P" & y).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Highlight the Unmatched Red in HFA
  For y = 2 To Worksheets(1).Range("A3000").End(xlUp).Row
    If Worksheets(1).Range("A" & y).Interior.ColorIndex = xlNone Then
      Worksheets(1).Range("A" & y).Interior.Color = rgbSalmon
      Worksheets(1).Range("B" & y).Interior.Color = rgbSalmon
      Worksheets(1).Range("C" & y).Interior.Color = rgbSalmon
      Worksheets(1).Range("D" & y).Interior.Color = rgbSalmon
      Worksheets(1).Range("E" & y).Interior.Color = rgbSalmon
      Worksheets(1).Range("F" & y).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Highlight Grid Descriptions in Oracle - Informational
  For j = 2 To Worksheets(3).Range("A3000").End(xlUp).Row
    If InStr(Worksheets(3).Range("D" & j), "GRD") <> 0 _
    Or InStr(Worksheets(3).Range("D" & j), "GRID") <> 0 _
    Or InStr(Worksheets(3).Range("D" & j), "MUNTIN") <> 0 Then
      Worksheets(3).Range("D" & j).Interior.Color = RGB(102, 102, 204)
    End If
  Next j
  
End Sub
 
  
'=================================================================================================================================================================================================
'FINAL CLEAN UP AND FORMATTING
'=================================================================================================================================================================================================
Sub CleanUp()

  'Format sheets to autofit and hide
  Worksheets(3).Columns("A:BG").Columns.AutoFit
  Worksheets(3).Columns("E:M").Hidden = True
  Worksheets(3).Columns("O").Hidden = True
  Worksheets(3).Columns("S:AV").Hidden = True
  Worksheets(3).Columns("R:BG").Hidden = True
  Worksheets(1).Columns("A:R").Columns.AutoFit
  Worksheets(2).Columns("A:AS").Columns.AutoFit
  Worksheets(4).Columns("A:AX").Columns.AutoFit
  
  Application.DisplayAlerts = False
  Worksheets(4).Delete
  Application.DisplayAlerts = True

  Application.ScreenUpdating = True
  
End Sub


'=================================================================================================================================================================================================
'Summary of BOM after comparision
'=================================================================================================================================================================================================

Sub C3POIsTheBest()

Application.ScreenUpdating = False

Dim startOracle As Single
Dim endOracle As Single
Dim startHfa As Single
Dim endHfa As Single

'startOracle = Timer()
Call MissingFromOracle
'endOracle = Timer()

'startHfa = Timer()
Call MissingFromHfa
'endHfa = Timer()



'Extra check to indicate whether boms match or not
Call DoBomsMatch

'uncomment to reformat (expand the columns)
Call ReFormat

Application.ScreenUpdating = True

'MsgBox ("Time taken to run Oracle code:" & endOracle - startOracle & " seconds" & vbNewLine & "Time taken to run HFA code: " & endHfa - startHfa & " seconds")


End Sub

'=================================================================================================================================================================================================
'Materials that is missing from Oracle DB
'=================================================================================================================================================================================================

Sub MissingFromOracle()

Dim ws As Worksheet
Set ws = Sheets(1)
Dim i As Long
i = 1
Dim cnt As Long
cnt = 6
Dim LastCol As Long
'LastCol = (Cells(Rows.Count, i).End(xlUp).Row) + 1
LastCol = (Worksheets(1).Range("A500").End(xlUp).Row) + 1
Debug.Print "length: "; LastCol

'Worksheets("Sheet1").Range("A1:D1").Copy Worksheets("Sheet2").Range("A1:D1")
'set the table headings
Worksheets("Validation").Range("BJ4").Value = "Missing from Oracle"
Worksheets(1).Range("E1").Copy Worksheets("Validation").Range("BJ5")
Worksheets(1).Range("F1").Copy Worksheets("Validation").Range("BK5")
Worksheets(1).Range("J1").Copy Worksheets("Validation").Range("BL5")
Worksheets(1).Range("K1").Copy Worksheets("Validation").Range("BM5")

'check to see if cell is red, if so then copy that cell over to the other sheet
Do Until i = LastCol
    If ws.Range("E" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("F" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("J" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("K" & i).Interior.Color = RGB(250, 128, 114) Then
    If ws.Range("E" & i).Interior.Color = RGB(250, 128, 114) Then
        ws.Range("E" & i).Copy Worksheets("Validation").Range("BJ" & cnt)
    End If
    If ws.Range("F" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.Range("E" & i).Copy Worksheets("Validation").Range("BJ" & cnt)
        End If
        ws.Range("F" & i).Copy Worksheets("Validation").Range("BK" & cnt)
    End If
    If ws.Range("J" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.Range("E" & i).Copy Worksheets("Validation").Range("BJ" & cnt)
        End If
        ws.Range("J" & i).Copy Worksheets("Validation").Range("BL" & cnt)
    End If
    If ws.Range("K" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.Range("E" & i).Copy Worksheets("Validation").Range("BJ" & cnt)
        End If
        ws.Range("K" & i).Copy Worksheets("Validation").Range("BM" & cnt)
    End If
    cnt = cnt + 1
    End If

i = i + 1

Loop
'Columns.AutoFit
'C3PO human cyborg relations - how may i help you
End Sub

'=================================================================================================================================================================================================
'Materials that is missing from HFA DB
'=================================================================================================================================================================================================

Sub MissingFromHfa()

Dim ws As Worksheet
Set ws = Sheets(3)
Dim i As Long
i = 1
Dim cnt As Long
cnt = 6
Dim LastCol As Long
'LastCol = (Cells(Rows.Count, i).End(xlUp).Row) + 1
LastCol = (Worksheets(3).Range("A3000").End(xlUp).Row) + 1
Debug.Print "length: "; LastCol

'Worksheets("Sheet1").Range("A1:D1").Copy Worksheets("Sheet2").Range("A1:D1")
'set the table headings
Worksheets("Validation").Range("BO4").Value = "Missing from HFA"
Worksheets(3).Range("C1").Copy Worksheets("Validation").Range("BO5")
Worksheets(3).Range("D1").Copy Worksheets("Validation").Range("BP5")
Worksheets(3).Range("P1").Copy Worksheets("Validation").Range("BQ5")
Worksheets(3).Range("Q1").Copy Worksheets("Validation").Range("BR5")

'check to see if cell is red, if so then copy that cell over to the other sheet
Do Until i = LastCol
    If ws.Range("C" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("D" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("P" & i).Interior.Color = RGB(250, 128, 114) Or ws.Range("Q" & i).Interior.Color = RGB(250, 128, 114) Then
    If ws.Range("C" & i).Interior.Color = RGB(250, 128, 114) Then
        'ws.Range("C" & i).Copy Range("BO" & cnt)
        ws.Range("C" & i).Copy Worksheets("Validation").Range("BO" & cnt)
        'ws.Range("C" & i).Value = Range("BO" & cnt).Value
    End If
    If ws.Range("D" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.Range("C" & i).Copy Worksheets("Validation").Range("BO" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("D" & i).Copy Range("BP" & cnt)
         ws.Range("D" & i).Copy Worksheets("Validation").Range("BP" & cnt)
        'ws.Range("D" & i).Value = Range("BP" & cnt).Value
    End If
    If ws.Range("P" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.Range("C" & i).Copy Worksheets("Validation").Range("BO" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("p" & i).Copy Range("BQ" & cnt)
        ws.Range("P" & i).Copy Worksheets("Validation").Range("BQ" & cnt)
        'ws.Range("P" & i).Value = Range("BQ" & cnt).Value
    End If
    If ws.Range("Q" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.Range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.Range("C" & i).Copy Worksheets("Validation").Range("BR" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("Q" & i).Copy Range("BR" & cnt)
        ws.Range("Q" & i).Copy Worksheets("Validation").Range("BR" & cnt)
        'ws.Range("Q" & i).Value = Range("BR" & cnt).Value
    End If
    cnt = cnt + 1
    End If

i = i + 1

Loop
'Columns.AutoFit
'C3PO human cyborg relations - how may i help you
End Sub

Sub ReFormat()
Worksheets(3).Columns("BJ:BR").Columns.AutoFit
'Worksheets(3).Columns("BO:BR").Columns.AutoFit
End Sub


'=================================================================================================================================================================================================
'Check to indicate whether BOM's are a match or not
'=================================================================================================================================================================================================

Sub DoBomsMatch()

Dim item As Variant

For Each item In Worksheets(1).Range("A2:R3000").Cells
    If item.Interior.Color = RGB(250, 128, 114) Then
        Worksheets("Validation").Range("BP1").Value = "ERRORS in HFA BOM"
        Worksheets("Validation").Range("BP1").Interior.Color = RGB(255, 0, 0)
        Exit For
    Else
        Worksheets("Validation").Range("BP1").Value = "NO Errors in HFA BOM"
        Worksheets("Validation").Range("BP1").Interior.Color = RGB(0, 255, 0)
    End If
Next

For Each item In Worksheets(3).Range("A2:BG3000").Cells
    If item.Interior.Color = RGB(250, 128, 114) Then
        Worksheets("Validation").Range("BS1").Value = "ERRORS in Oracle BOM"
        Worksheets("Validation").Range("BS1").Interior.Color = RGB(255, 0, 0)
        Exit For
    Else
        Worksheets("Validation").Range("BS1").Value = "NO Errors in Oracle BOM"
        Worksheets("Validation").Range("BS1").Interior.Color = RGB(0, 255, 0)
    End If
Next

End Sub









