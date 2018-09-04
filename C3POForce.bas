Sub Main()

Call LightForce
Call C3P0IsTheBest

End Sub

'=================================================================================================================================================================================================
'Check if worksheets are present and that one is hfa and the other is oracle. If not - exit out of the program and display error message to user
'=================================================================================================================================================================================================
Sub CheckWorksheets()

Dim bSheetIsEmpty As Boolean
Dim ur As range
Dim cell As range
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    Set ur = ws.UsedRange
    If ur.count = 1 Then
        bSheetIsEmpty = Len(ur) = 0
    Else
        Set cell = ur.Cells.Find("*", ur(1), -4123&, 2&, 2&, 0, 0)
        bSheetIsEmpty = cell Is Nothing
    End If
    'Debug.Print ws.Name, bSheetIsEmpty
    If bSheetIsEmpty = True Then
        MsgBox ws.Name & " is empty"
        Exit Sub
    End If
        
Next

Dim i As Long
Dim wsExists As Boolean
wsExists = False
Dim wsExists2 As Boolean
wsExists2 = False

For i = 1 To Worksheets.count
    'Debug.Print ("iter pos: " & i)
    'before running light force - must have only two spreedsheets
    If i > 2 And Worksheets(i).Name Like "Validation" Then
        'MsgBox ("Error: make sure there are only two spreadsheets, 1 - HFA BOM, 2 - Oracle BOM. Please remove the extra worksheets. Program execution TERMINATED.")
        'if users wants to rerun lightforce on a USED workbook then call RerunCheck()
        Call RerunCheck
        Exit Sub
    End If
    If Worksheets(i).Name Like "MIL5WIFX*" Then
        wsExists = True
        'MsgBox ("Success: hfa worksheet present")
    End If
    If Worksheets(i).Name Like "fnd_gfm_*" Then
        wsExists2 = True
        'MsgBox ("Success: Oracle worksheet present")
    End If
Next i

'Debug.Print ("1: " & wsExists)
'Debug.Print ("2: " & wsExists2)

If wsExists = False And wsExists2 = False Then
    MsgBox ("ERROR: Both sheets are either named incorrect or in incorrect format. Oracle BOM is not present (  worksheet name must be named as such: fng_gfm_XXXXXXX  ). HFA bom is not present (  worksheet name must be named as such: MIL5WIFX(XXX)  ). Program execution TERMINATED.")
    Exit Sub
End If
If Not wsExists Then
    MsgBox ("ERROR: HFA bom is not present (  worksheet name must be named as such: MIL5WIFX(XXX)  ). Program execution TERMINATED.")
    Exit Sub
End If
If Not wsExists2 Then
    MsgBox ("ERROR: Oracle BOM is not present (  worksheet name must be named as such: fng_gfm_XXXXXXX  ). Program execution TERMINATED.")
    Exit Sub
End If

'if passes all checks, execute main
Call Main


End Sub

'=================================================================================================================================================================================================
'If users wants to rerun C3POForce - to avoid previous crash, check if the sheets(1 and 3) are colored, if so then lightforce was already execute - rerun C3POIsTheBest
'=================================================================================================================================================================================================
Sub RerunCheck()

'MsgBox ("rerun test")
Dim isColored As Boolean
isColored = False

For i = 1 To Worksheets.count
    Debug.Print (i)
    If range("A2").Cells.Interior.ColorIndex > 0 Then
    'If range("A2", range("A2").End(xlDown)).Cells.Interior.ColorIndex > 0 Then
        isColored = True
    Else
        isColored = False
    End If
    'skip second sheet
    i = i + 1
Next i


If isColored = True Then
    Call C3POIsTheBest
End If

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
  Worksheets(3).range("B1").EntireColumn.Insert
  Worksheets(3).range("B1").Value = "Validation"
  Sheets.Add After:=Worksheets(3)
  Worksheets(4).Name = "Cuts"
  
  
  'Set up for loop
  Dim glassCount As Integer, hFACutIndex As Integer, cutFlatten As Integer, j As Integer, i As Integer
  Dim hFACutLower As Double, hFACutUpper As Double
  
  glassCount = 1
  hFACutIndex = 1

  Worksheets(3).Columns(17).ClearContents
  Worksheets(3).Columns(18).ClearContents
  Worksheets(3).range("Q1").Value = "HFA Cut Length"
  
  'OUTER LOOP HFA
  For j = 2 To Worksheets(1).range("C3000").End(xlUp).Row
    'HFA
    'Highlight Glass
    If InStr(Worksheets(1).range("E" & j), "GT") = 1 _
    Or InStr(Worksheets(1).range("E" & j), "GA") = 1 Then
      glassCount = glassCount + 1
      Worksheets(1).range("A" & j).Interior.Color = rgbOrange
      Worksheets(1).range("B" & j).Interior.Color = rgbOrange
      Worksheets(1).range("C" & j).Interior.Color = rgbOrange
      Worksheets(1).range("D" & j).Interior.Color = rgbOrange
      Worksheets(1).range("E" & j).Interior.Color = rgbOrange
      Worksheets(3).range("AZ" & glassCount).Value = Worksheets(1).range("J" & j).Value
      Worksheets(3).range("BA" & glassCount).Value = Worksheets(1).range("K" & j).Value
      'Track Position of HFA for Highlighting Later
      Worksheets(3).range("BG1").Value = "HFA BOM line"
      Worksheets(3).range("BG" & glassCount).Value = j
    End If
    'HFA
    'Highlight Frame and other parts that are not relivant and Spacer
    If InStr(Worksheets(1).range("E" & j), "VIG") = 1 _
    Or InStr(Worksheets(1).range("E" & j), "VP") = 1 _
    Or InStr(Worksheets(1).range("E" & j), "FR") = 1 _
    Or InStr(Worksheets(1).range("E" & j), "FP") = 1 _
    Or InStr(Worksheets(1).range("E" & j), "WNIG") <> 0 _
    Or InStr(Worksheets(1).range("E" & j), "PAINTING") <> 0 _
    Or (InStr(Worksheets(1).range("Q" & j), "SPACER") <> 0 And Worksheets(1).range("H" & j).Value = "LI") Then
      Worksheets(1).range("A" & j).Interior.Color = rgbGrey
      Worksheets(1).range("B" & j).Interior.Color = rgbGrey
      Worksheets(1).range("C" & j).Interior.Color = rgbGrey
      Worksheets(1).range("D" & j).Interior.Color = rgbGrey
      Worksheets(1).range("E" & j).Interior.Color = rgbGrey
      Worksheets(1).range("F" & j).Interior.Color = rgbGrey
      Worksheets(1).range("Q" & j).Interior.Color = rgbGrey
      If InStr(Worksheets(1).range("Q" & j), "SPACER") <> 0 Then
        Worksheets(1).range("J" & j).Interior.Color = rgbGrey
      End If
    End If
    
    'HFA Minor Parts
    If InStr(Worksheets(1).range("Q" & j), "CLIP") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "BREATHER TUBE") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "BUTYL") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), " DSE ") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "CLILP") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "DESICCANT") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "GASKET") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "TAPE") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "SASH STOP") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "SETTING BLOCK") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "CAULKING") <> 0 _
    Or InStr(Worksheets(1).range("E" & j), "WS") = 1 _
    Or InStr(Worksheets(1).range("Q" & j), "CVR") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "COVER") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "STRIKE") <> 0 Then
      Worksheets(1).range("A" & j).Interior.Color = rgbGrey
      Worksheets(1).range("B" & j).Interior.Color = rgbGrey
      Worksheets(1).range("C" & j).Interior.Color = rgbGrey
      Worksheets(1).range("D" & j).Interior.Color = rgbGrey
      Worksheets(1).range("E" & j).Interior.Color = rgbGrey
      Worksheets(1).range("F" & j).Interior.Color = rgbGrey
      Worksheets(1).range("Q" & j).Interior.Color = rgbGrey
    End If
    
    'Highlight Grid Descriptions
    If InStr(Worksheets(1).range("Q" & j), "GRD") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "GRID") <> 0 _
    Or InStr(Worksheets(1).range("Q" & j), "MUNTIN") <> 0 Then
      Worksheets(1).range("Q" & j).Interior.Color = RGB(102, 102, 204)
    End If
    
    'HFA
    'NEW Add Cut parts to 4th Worksheet "Cuts"
    If (Worksheets(1).range("A" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).range("A" & j).Interior.Color <> rgbOrange _
    And Worksheets(1).range("J" & j).Value > 1 _
    And Worksheets(1).range("H" & j).Value = "LI" _
    And InStr(Worksheets(1).range("Q" & j), "SPACER") = 0) _
 _
    Or (Worksheets(1).range("A" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).range("A" & j).Interior.Color <> rgbOrange _
    And Worksheets(1).range("H" & j).Value = "EA" _
    And InStr(Worksheets(1).range("Q" & j), "PRECUT V") <> 0) Then
        For cutFlatten = Worksheets(1).range("F" & j).Value To 1 Step -1
          Worksheets(4).range("A" & hFACutIndex).Value = j
          Worksheets(4).range("B" & hFACutIndex).Value = Worksheets(1).range("E" & j).Value
          Worksheets(4).range("B" & hFACutIndex).Interior.Color = rgbRed
          hFACutLower = Val(Worksheets(1).range("J" & j).Value) - 0.0625
          hFACutUpper = Val(Worksheets(1).range("J" & j).Value) + 0.0625
          Worksheets(4).range("C" & hFACutIndex).Value = hFACutLower
          Worksheets(4).range("D" & hFACutIndex).Value = Worksheets(1).range("J" & j).Value
          Worksheets(4).range("E" & hFACutIndex).Value = hFACutUpper
          Worksheets(4).range("F" & hFACutIndex).Value = Worksheets(1).range("F" & j).Value
          hFACutIndex = hFACutIndex + 1
        Next cutFlatten
    End If
     
'------------------------------------------------------------------------------------
    
    'INNER LOOP FOR ORACLE PARTS
    For i = 2 To Worksheets(3).range("C3000").End(xlUp).Row
      'Oracle
      'Highlight All UOM EA And LI
      Worksheets(3).range("N" & i).Interior.Color = rgbGrey
      'Oracle
      'Grey out any lines that will not be matched like Sublines, IGU, PANEL, LABEL, FRAME, and GLASS
      'Or (InStr(Worksheets(3).Range("C" & i), "*") <> 0 And InStr(Worksheets(3).Range("C" & i), "X-") = 0)
      If Worksheets(3).range("N" & i).Value = "LI" _
      Or InStr(1, Worksheets(3).range("C" & i), "PANEL") _
      Or InStr(1, Worksheets(3).range("C" & i), "LABEL") _
      Or InStr(1, Worksheets(3).range("C" & i), "FRAME") _
      Or InStr(1, Worksheets(3).range("C" & i), "X-GT") _
      Or InStr(1, Worksheets(3).range("C" & i), "X-GA") _
      Or InStr(1, Worksheets(3).range("C" & i), "GA") _
      Or InStr(1, Worksheets(3).range("C" & i), "GT") Then
        Worksheets(3).range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).range("P" & i).Interior.Color = rgbGrey
      End If
      
      'Oracle
      'Grey out Precuts since they are 'EA'
      If InStr(Worksheets(3).range("D" & i), "PRECUT V") <> 0 Then
        Worksheets(3).range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).range("P" & i).Interior.Color = rgbGrey
      End If
      
      'Oracle Minor Parts and Spacer LI
      If InStr(Worksheets(3).range("D" & i), "CLIP") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "BREATHER TUBE") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "WEEP") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "BUTYL") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "CVR") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "Cover") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), " DSE ") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "DESICCANT") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "GASKET") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "TAPE") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "SASH STOP") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "BLOCK") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "CAULKING") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "STRIKE") <> 0 _
      Or InStr(Worksheets(3).range("D" & i), "ARGON") <> 0 _
      Or (InStr(Worksheets(3).range("D" & i), "SPACER") <> 0 And Worksheets(3).range("N" & i).Value = "LI") Then
        Worksheets(3).range("A" & i).Interior.Color = rgbGrey
        Worksheets(3).range("B" & i).Interior.Color = rgbGrey
        Worksheets(3).range("C" & i).Interior.Color = rgbGrey
        Worksheets(3).range("D" & i).Interior.Color = rgbGrey
        Worksheets(3).range("P" & i).Interior.Color = rgbGrey
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
  Worksheets(3).range("AW1").Value = "Oracle GLASS W"
  Worksheets(3).range("AX1").Value = "Oracle GLASS H"
  
  Dim c As Integer, widthStopCount As Integer
  Dim glassWidthStr As String, glassHeightStr As String
  Dim widthStop As Boolean
  Dim glassWidth As Double, glassHeight As Double
  
  glassCount = 1
  
  For i = 2 To Worksheets(3).range("C3000").End(xlUp).Row
    'Parse out Width And Height from IGU line And Print to new Column
    If InStr(1, Worksheets(3).range("C" & i), "IGU") Then
      glassCount = glassCount + 1
      glassWidthStr = ""
      glassHeightStr = ""
      widthStop = False
      widthStopCount = 0
          
      'Look in Description to get width and height
      For c = 1 To Len(Worksheets(3).range("D" & i).Value)
        Dim currentChar As String
        currentChar = Mid(Worksheets(3).range("D" & i).Value, c, 1)
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
            Worksheets(3).range("AW" & glassCount).Value = glassWidth
            Worksheets(3).range("AX" & glassCount).Value = glassHeight
            Worksheets(3).range("AY1").Value = "Oracle BOM line"
            Worksheets(3).range("AY" & glassCount).Value = i
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
  Worksheets(3).range("AZ1").Value = "HFA GLASS W"
  Worksheets(3).range("BA1").Value = "HFA GLASS H"
  
  Dim lastrowHFAGlass As Long
  Dim position As Integer
  
  lastrowHFAGlass = Worksheets(3).range("AZ50").End(xlUp).Row
  For i = 2 To lastrowHFAGlass
    For j = 3 To lastrowHFAGlass
      If i <> j _
      And Worksheets(3).range("AZ" & i).Value = Worksheets(3).range("AZ" & j).Value _
      And Worksheets(3).range("BA" & i).Value = Worksheets(3).range("BA" & j).Value _
      And Worksheets(3).range("BB" & j).Value <> "flagged" _
      And IsEmpty(Worksheets(3).range("AY" & j).Value) Then
        Worksheets(3).range("BB" & i).Value = "flagged"
        Worksheets(3).range("AZ" & j).Value = "Duplicate"
        Worksheets(3).range("BA" & j).Value = "Duplicate"
        'Grey out HFA sheet measurement for Duplicate
        position = Worksheets(3).range("BG" & j).Value
        Worksheets(1).range("J" & position).Interior.Color = rgbGrey
        Worksheets(1).range("K" & position).Interior.Color = rgbGrey
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
  
  lastrowOracleGlass = Worksheets(3).range("AW50").End(xlUp).Row
   
  For j = 2 To lastrowHFAGlass
    hfaWidthCut = Worksheets(3).range("AZ" & j).Value
    hFAWidthCutLower = Val(Worksheets(3).range("AZ" & j).Value) - 0.0625
    hFAWidthCutUpper = Val(Worksheets(3).range("AZ" & j).Value) + 0.0625
        
    hfaHeightCut = Worksheets(3).range("BA" & j).Value
    hFAHeightCutLower = Val(Worksheets(3).range("BA" & j).Value) - 0.0625
    hFAHeightCutUpper = Val(Worksheets(3).range("BA" & j).Value) + 0.0625
    
    Worksheets(3).range("BC" & j).Value = hFAWidthCutLower
    Worksheets(3).range("BD" & j).Value = hFAWidthCutUpper
    Worksheets(3).range("BE" & j).Value = hFAHeightCutLower
    Worksheets(3).range("BF" & j).Value = hFAHeightCutUpper
  
    For i = 2 To lastrowOracleGlass
      'See if both W and H match
      If Worksheets(3).range("AW" & i).Value >= hFAWidthCutLower _
      And Worksheets(3).range("AW" & i).Value <= hFAWidthCutUpper _
      And Worksheets(3).range("AX" & i).Value >= hFAHeightCutLower _
      And Worksheets(3).range("AX" & i).Value <= hFAHeightCutUpper _
      And Worksheets(3).range("AY" & i).Interior.Color <> rgbGreen Then
        position = Worksheets(3).range("AY" & i).Value
        Worksheets(3).range("AY" & i).Interior.Color = rgbGreen
        Worksheets(3).range("D" & position).Interior.Color = rgbGreen
        Worksheets(3).range("A" & position).Interior.Color = rgbAqua
        Worksheets(3).range("B" & position).Interior.Color = rgbAqua
        Worksheets(3).range("C" & position).Interior.Color = rgbAqua
        Worksheets(3).range("P" & position).Interior.Color = rgbAqua
        Worksheets(3).range("B" & position).Value = "Correct Size"
        'Input match for glass in Q if there is a match
        Worksheets(3).range("Q" & position).Value = hfaWidthCut + " X " + hfaHeightCut
        Worksheets(3).range("Q" & position).Interior.Color = rgbGreen
        'Highlight HFA Page with correct glass
        position = Worksheets(3).range("BG" & i).Value
        Worksheets(1).range("J" & position).Interior.Color = rgbGreen
        Worksheets(1).range("K" & position).Interior.Color = rgbGreen
        Exit For
      End If
    Next i
  Next j
  
  'Glass Cleanup highlight bad glass
  'HFA
  For j = 2 To Worksheets(1).range("A3000").End(xlUp).Row
    If Worksheets(1).range("E" & j).Interior.Color = rgbOrange _
    And Worksheets(1).range("J" & j).Interior.Color <> rgbGrey _
    And Worksheets(1).range("J" & j).Interior.Color <> rgbGreen Then
       Worksheets(1).range("J" & j).Interior.Color = rgbSalmon
       Worksheets(1).range("K" & j).Interior.Color = rgbSalmon
    End If
  Next j
  
  For j = 2 To Worksheets(3).range("C3000").End(xlUp).Row
    If InStr(1, Worksheets(3).range("C" & j), "IGU") <> 0 _
    And Worksheets(3).range("C" & j).Interior.Color <> rgbAqua Then
       Worksheets(3).range("A" & j).Interior.Color = rgbSalmon
       Worksheets(3).range("B" & j).Interior.Color = rgbSalmon
       Worksheets(3).range("C" & j).Interior.Color = rgbSalmon
       Worksheets(3).range("D" & j).Interior.Color = rgbSalmon
       Worksheets(3).range("P" & j).Interior.Color = rgbSalmon
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
  
  For j = 2 To Worksheets(1).range("C3000").End(xlUp).Row
    If Worksheets(1).range("H" & j).Value = "EA" _
    And Worksheets(1).range("K" & j).Value <> 0 Then
      For screenFlatten = Worksheets(1).range("F" & j).Value To 1 Step -1
        Worksheets(4).range("AD" & screenCount).Value = j
        Worksheets(4).range("AE" & screenCount).Value = Worksheets(1).range("E" & j).Value
        Worksheets(4).range("AF" & screenCount).Value = Worksheets(1).range("J" & j).Value - 0.0625
        Worksheets(4).range("AG" & screenCount).Value = Worksheets(1).range("J" & j).Value
        Worksheets(4).range("AH" & screenCount).Value = Worksheets(1).range("J" & j).Value + 0.0625
        Worksheets(4).range("AI" & screenCount).Value = Worksheets(1).range("K" & j).Value - 0.0625
        Worksheets(4).range("AJ" & screenCount).Value = Worksheets(1).range("K" & j).Value
        Worksheets(4).range("AK" & screenCount).Value = Worksheets(1).range("K" & j).Value + 0.0625
        screenCount = screenCount + 1
      Next screenFlatten
    End If
  Next j
  
  'Oracle
  screenCount = 1
  For j = 2 To Worksheets(3).range("C3000").End(xlUp).Row
    If InStr(1, Worksheets(3).range("D" & j).Value, "Screen,") Then
      Worksheets(4).range("AP" & screenCount).Value = j
      Worksheets(4).range("AQ" & screenCount).Value = Worksheets(3).range("C" & j).Value
      Worksheets(4).range("AR" & screenCount).Value = Worksheets(3).range("D" & j).Value
      'Width
      xMarker = InStr(1, Worksheets(4).range("AR" & screenCount).Value, "x")
      'Hundreds Place
      If IsNumeric(Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker - 9, 1)) Then
        Worksheets(4).range("AS" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker - 9, 8)
        Worksheets(4).range("AS" & screenCount).Interior.Color = RGB(102, 102, 204)
      'Tens Place
      ElseIf IsNumeric(Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker - 8, 1)) Then
        Worksheets(4).range("AS" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker - 8, 7)
        Worksheets(4).range("AS" & screenCount).Interior.Color = rgbGreen
      'Ones Place
      Else
        Worksheets(4).range("AS" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker - 7, 6)
        Worksheets(4).range("AS" & screenCount).Interior.Color = rgbGold
      End If
      'Height
      'Hundreds Place
      If IsNumeric(Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker + 9, 1)) Then
        Worksheets(4).range("AT" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker + 2, 8)
        Worksheets(4).range("AT" & screenCount).Interior.Color = RGB(102, 102, 204)
      'Tens Place
      ElseIf IsNumeric(Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker + 8, 1)) Then
        Worksheets(4).range("AT" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker + 2, 7)
        Worksheets(4).range("AT" & screenCount).Interior.Color = rgbGreen
      'Ones Place
      Else
        Worksheets(4).range("AT" & screenCount).Value = Mid(Worksheets(4).range("AR" & screenCount).Value, xMarker + 2, 6)
        Worksheets(4).range("AT" & screenCount).Interior.Color = rgbGold
      End If
      screenCount = screenCount + 1
    End If
  Next j
  
  'Perfect Match
  'Compare Oracle to HFA
  screenCount = 1
  If Worksheets(4).range("AQ1").Value <> "" And Worksheets(4).range("AE1").Value <> "" Then
    For j = 1 To Worksheets(4).range("AQ50").End(xlUp).Row
      For i = 1 To Worksheets(4).range("AE50").End(xlUp).Row
        If InStr(1, Worksheets(4).range("AQ" & j).Value, Worksheets(4).range("AE" & i).Value) <> 0 _
        And Worksheets(4).range("AS" & j).Value >= Worksheets(4).range("AF" & i).Value _
        And Worksheets(4).range("AS" & j).Value <= Worksheets(4).range("AH" & i).Value _
        And Worksheets(4).range("AT" & j).Value >= Worksheets(4).range("AI" & i).Value _
        And Worksheets(4).range("AT" & j).Value <= Worksheets(4).range("AK" & i).Value _
        And Worksheets(4).range("AQ" & j).Interior.Color <> rgbGreen _
        And Worksheets(4).range("AE" & i).Interior.Color <> rgbGreen Then
          'Page 4
          Worksheets(4).range("AQ" & j).Interior.Color = rgbGreen
          Worksheets(4).range("AE" & i).Interior.Color = rgbGreen
          'Page 3
          Worksheets(3).range("A" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).range("B" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).range("C" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).range("D" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).range("P" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).range("Q" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).range("Q" & Worksheets(4).range("AP" & j).Value).Value = Worksheets(4).range("AG" & i).Value & " X " & Worksheets(4).range("AJ" & i).Value
          'Page 1
          Worksheets(1).range("A" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("B" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("C" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("D" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("E" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("F" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("J" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbGreen
          Worksheets(1).range("K" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbGreen
        End If
      Next i
    Next j
    
    'Match If Cuts are equal but part is not
    'Compare Oracle to HFA
    screenCount = 1
    For j = 1 To Worksheets(4).range("AQ50").End(xlUp).Row
      For i = 1 To Worksheets(4).range("AE50").End(xlUp).Row
        If Worksheets(4).range("AS" & j).Value >= Worksheets(4).range("AF" & i).Value _
        And Worksheets(4).range("AS" & j).Value <= Worksheets(4).range("AH" & i).Value _
        And Worksheets(4).range("AT" & j).Value >= Worksheets(4).range("AI" & i).Value _
        And Worksheets(4).range("AT" & j).Value <= Worksheets(4).range("AK" & i).Value _
        And Worksheets(4).range("AQ" & j).Interior.Color <> rgbGreen _
        And Worksheets(4).range("AE" & i).Interior.Color <> rgbGreen Then
          'Page 4
          Worksheets(4).range("AQ" & j).Interior.Color = rgbGreen
          Worksheets(4).range("AE" & i).Interior.Color = rgbGreen
          'Page 3
          Worksheets(3).range("A" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).range("B" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).range("C" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
          Worksheets(3).range("D" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).range("P" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbAqua
          Worksheets(3).range("Q" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbGreen
          Worksheets(3).range("Q" & Worksheets(4).range("AP" & j).Value).Value = Worksheets(4).range("AG" & i).Value & " X " & Worksheets(4).range("AJ" & i).Value
          'Page 1
          Worksheets(1).range("A" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).range("B" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).range("C" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).range("D" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).range("E" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbSalmon
          Worksheets(1).range("F" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbAqua
          Worksheets(1).range("J" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbGreen
          Worksheets(1).range("K" & Worksheets(4).range("AD" & i).Value).Interior.Color = rgbGreen
        End If
      Next i
    Next j
    
    'CleanUp
    'Oracle
    For j = 1 To Worksheets(4).range("AQ50").End(xlUp).Row
      If Worksheets(4).range("AQ" & j).Interior.Color <> rgbGreen Then
        Worksheets(3).range("A" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).range("B" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).range("C" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).range("D" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
        Worksheets(3).range("P" & Worksheets(4).range("AP" & j).Value).Interior.Color = rgbSalmon
      End If
    Next j
    'HFA
    For j = 1 To Worksheets(4).range("AE50").End(xlUp).Row
      If Worksheets(4).range("AE" & j).Interior.Color <> rgbGreen Then
        Worksheets(1).range("A" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("B" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("C" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("D" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("E" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("F" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("J" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
        Worksheets(1).range("K" & Worksheets(4).range("AD" & j).Value).Interior.Color = rgbSalmon
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
  For i = 2 To Worksheets(3).range("C3000").End(xlUp).Row
      If Worksheets(3).range("A" & i).Interior.Color <> rgbGrey _
      And InStr(Worksheets(3).range("C" & i), "X-") <> 0 Then
        hasRun = False
        startIndex = 0
        strIndexCount = 0
        buildingString = ""
        For b = 1 To Len(Worksheets(3).range("D" & i).Value)
          Dim currentDChar As String
          currentDChar = Mid(Worksheets(3).range("D" & i).Value, b, 1)
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
        For j = (InStr(1, Worksheets(3).range("C" & i).Value, "X-") + 2) To Len(Worksheets(3).range("C" & i).Value)
          If Mid(Worksheets(3).range("C" & i).Value, j, 1) = "*" Then
            Exit For
          End If
          starSlice = starSlice + Mid(Worksheets(3).range("C" & i).Value, j, 1)
        Next j
        
        descriptionStr = Worksheets(3).range("D" & i).Value
        oracleCutStr = Mid(descriptionStr, startIndex, strIndexCount + 1)
        oracleCut = Val(oracleCutStr)
    
        For oracleFlatten = Worksheets(3).range("P" & i).Value To 1 Step -1
          Worksheets(4).range("J" & oracleCutIndex).Value = i
          Worksheets(4).range("K" & oracleCutIndex).Value = Worksheets(3).range("C" & i).Value
          Worksheets(4).range("L" & oracleCutIndex).Value = oracleCut
          Worksheets(4).range("M" & oracleCutIndex).Value = Worksheets(3).range("P" & i).Value
          Worksheets(4).range("N" & oracleCutIndex).Value = starSlice
          oracleCutIndex = oracleCutIndex + 1
        Next oracleFlatten
      End If
  Next i
  
 'Sort Oracle list A-Z
  Worksheets(4).range("J1:N100").Sort Key1:=range("N1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
  
  'PERFECT MATCH - GREEN
  'Oracle to HFA comparison match
  'All HFA column starts out RED
  For y = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    match = False
    For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
      If Worksheets(4).range("N" & y).Value = Worksheets(4).range("B" & z).Value _
      And Worksheets(4).range("B" & z).Interior.Color <> rgbGreen _
      And ((Worksheets(4).range("L" & y).Value >= Worksheets(4).range("C" & z).Value And Worksheets(4).range("L" & y).Value <= Worksheets(4).range("E" & z).Value) Or Worksheets(4).range("D" & z).Value = 1) Then
        Worksheets(4).range("N" & y).Interior.Color = rgbGreen
        Worksheets(4).range("B" & z).Interior.Color = rgbGreen
        match = True
        Exit For
      End If
    Next z
  Next y
  
 'BAD CUTS - BLUE
 'Loop to compare unmatched Oracle column to HFA. Identify items that match but the lengths do not.
 'This will catch any bad cuts and before addressing quantity
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.Color <> rgbGreen _
    And Worksheets(4).range("N" & z).Interior.Color <> rgbGold Then
      For y = 1 To Worksheets(4).range("B3000").End(xlUp).Row
        If Worksheets(4).range("B" & y).Value = Worksheets(4).range("N" & z).Value _
        And Worksheets(4).range("B" & y).Interior.Color = rgbRed Then
          Worksheets(4).range("N" & z).Interior.Color = RGB(102, 102, 204)
          Worksheets(4).range("B" & y).Interior.Color = RGB(102, 102, 204)
          Exit For
        End If
      Next y
    End If
  Next z
  
  'BAD QUANTITY - GOLD
  'Check "HFA" Column for ID's that are both green and blank due to unmatched "HFA" lines
  'Partial approved lines due to quantity off
  For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
    If Worksheets(4).range("B" & z).Interior.Color <> rgbGreen Then
      For y = 1 To Worksheets(4).range("B3000").End(xlUp).Row
        If Worksheets(4).range("A" & y).Value = Worksheets(4).range("A" & z).Value _
        And Worksheets(4).range("B" & y).Interior.Color = rgbGreen Then
          'Change from blank to gold
          Worksheets(4).range("B" & z).Interior.Color = rgbGold
          Exit For
        End If
      Next y
    End If
  Next z
  'Check "Oracle" Column for ID's that are both green and blank due to unmatched "Oracle" lines
  'Partial approved lines due to quantity off
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.Color <> rgbGreen Then
      For y = 1 To Worksheets(4).range("N3000").End(xlUp).Row
        If Worksheets(4).range("J" & y).Value = Worksheets(4).range("J" & z).Value _
        And Worksheets(4).range("N" & y).Interior.Color = rgbGreen Then
          'Change from blank to gold
          Worksheets(4).range("N" & z).Interior.Color = rgbGold
          Exit For
        End If
      Next y
    End If
  Next z
 
'------------------------------------------------------------------------------------------------
  'ORACLE
  'Hightlight Validation page with Green
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.Color = rgbGreen Then
      Worksheets(3).range("D" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbGreen
      Worksheets(3).range("A" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("B" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("C" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("P" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Blue
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.Color = RGB(102, 102, 204) Then
      Worksheets(3).range("D" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("A" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("B" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("C" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("P" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Gold
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.Color = rgbGold Then
      Worksheets(3).range("A" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("B" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("C" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbAqua
      Worksheets(3).range("P" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      If Worksheets(3).range("D" & Worksheets(4).range("J" & z).Value).Interior.Color <> rgbGreen Then
        Worksheets(3).range("D" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      End If
    End If
  Next z
  
  'Hightlight Validation page with Blank
  For z = 1 To Worksheets(4).range("N3000").End(xlUp).Row
    If Worksheets(4).range("N" & z).Interior.ColorIndex = xlNone Then
      Worksheets(3).range("A" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("B" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("C" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("D" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("P" & Worksheets(4).range("J" & z).Value).Interior.Color = rgbSalmon
    
    End If
  Next z
  
'----------------------------------------------------------------------------------------------
  'HFA
  'Hightlight Validation page with Green
  For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
    If Worksheets(4).range("B" & z).Interior.Color = rgbGreen Then
      Worksheets(1).range("J" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbGreen
      Worksheets(1).range("A" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("B" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("C" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("D" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("E" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("F" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Blue
  For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
    If Worksheets(4).range("B" & z).Interior.Color = RGB(102, 102, 204) Then
      Worksheets(1).range("J" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("A" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("B" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("C" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("D" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("E" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("F" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
    End If
  Next z
  
  'Hightlight Validation page with Gold
  For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
    If Worksheets(4).range("B" & z).Interior.Color = rgbGold Then
      Worksheets(1).range("A" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("B" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("C" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("D" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("E" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbAqua
      Worksheets(1).range("F" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      If Worksheets(1).range("D" & Worksheets(4).range("A" & z).Value).Interior.Color <> rgbGreen Then
        Worksheets(1).range("J" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      End If
    End If
  Next z
  
  'Hightlight Validation page with Red
  For z = 1 To Worksheets(4).range("B3000").End(xlUp).Row
    If Worksheets(4).range("B" & z).Interior.Color = rgbRed Then
      Worksheets(1).range("A" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("B" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("C" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("D" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("E" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("F" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("J" & Worksheets(4).range("A" & z).Value).Interior.Color = rgbSalmon
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
  For z = 2 To Worksheets(3).range("A3000").End(xlUp).Row
    If Worksheets(3).range("C" & z).Interior.ColorIndex = xlNone Then
      'HFA
      For y = 2 To Worksheets(1).range("A3000").End(xlUp).Row
        If Worksheets(1).range("E" & y).Interior.ColorIndex = xlNone _
        And InStr(1, Worksheets(3).range("C" & z).Value, Worksheets(1).range("E" & y).Value) <> 0 _
        And InStr(1, Worksheets(3).range("P" & z).Value, Worksheets(1).range("F" & y).Value) <> 0 Then
        'Or InStr(1, Worksheets(1).Range("F" & y).Value, ".") <> 0) Then
          Worksheets(3).range("A" & z).Interior.Color = rgbAqua
          Worksheets(3).range("B" & z).Interior.Color = rgbAqua
          Worksheets(3).range("C" & z).Interior.Color = rgbAqua
          Worksheets(3).range("D" & z).Interior.Color = rgbAqua
          'If InStr(1, Worksheets(1).Range("F" & y).Value, ".") Then
            'Worksheets(3).Range("P" & z).Interior.Color = rgbGrey
          'Else
            Worksheets(3).range("P" & z).Interior.Color = rgbAqua
          'End If
          
          'HFA
          Worksheets(1).range("A" & y).Interior.Color = rgbAqua
          Worksheets(1).range("B" & y).Interior.Color = rgbAqua
          Worksheets(1).range("C" & y).Interior.Color = rgbAqua
          Worksheets(1).range("D" & y).Interior.Color = rgbAqua
          Worksheets(1).range("E" & y).Interior.Color = rgbAqua
          'If InStr(1, Worksheets(1).Range("F" & y).Value, ".") Then
            'Worksheets(1).Range("F" & y).Interior.Color = rgbGrey
          'Else
            Worksheets(1).range("F" & y).Interior.Color = rgbAqua
          'End If
          Exit For
        End If
      Next y
    End If
  Next z
  
  'HFA Flatten Remaining Parts
  'EA
  For y = 2 To Worksheets(1).range("A3000").End(xlUp).Row
    If Worksheets(1).range("C" & y).Interior.ColorIndex = xlNone _
    And InStr(1, Worksheets(1).range("H" & y).Value, "EA") = 1 Then
      For i = Worksheets(1).range("F" & y).Value To 1 Step -1
        Worksheets(4).range("R" & finalFlatten).Value = y
        Worksheets(4).range("S" & finalFlatten).Value = Worksheets(1).range("E" & y).Value
        Worksheets(4).range("T" & finalFlatten).Value = Worksheets(1).range("F" & y).Value
        finalFlatten = finalFlatten + 1
      Next i
    End If
  Next y
  
  'Oracle Flatten Remaining Parts
  'EA
  finalFlatten = 1
  For y = 2 To Worksheets(3).range("A3000").End(xlUp).Row
    If Worksheets(3).range("A" & y).Interior.ColorIndex = xlNone _
    And InStr(1, Worksheets(3).range("N" & y).Value, "EA") = 1 Then
      For i = Worksheets(3).range("P" & y).Value To 1 Step -1
        Worksheets(4).range("X" & finalFlatten).Value = y
        Worksheets(4).range("Y" & finalFlatten).Value = Worksheets(3).range("C" & y).Value
        Worksheets(4).range("Z" & finalFlatten).Value = Worksheets(3).range("P" & y).Value
        finalFlatten = finalFlatten + 1
      Next i
    End If
  Next y
  
  'Sort Oracle list A-Z
  Worksheets(4).range("X1:Z100").Sort Key1:=range("Y1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
     
  'Oracle compare to HFA
  For y = 1 To Worksheets(4).range("Y3000").End(xlUp).Row
    For z = 1 To Worksheets(4).range("S3000").End(xlUp).Row
      If InStr(Worksheets(4).range("Y" & y).Value, Worksheets(4).range("S" & z).Value) <> 0 _
      And Worksheets(4).range("Y" & y).Interior.Color <> rgbOrange _
      And Worksheets(4).range("S" & z).Interior.Color <> rgbOrange Then
        Worksheets(4).range("Y" & y).Interior.Color = rgbOrange
        Worksheets(4).range("S" & z).Interior.Color = rgbOrange
        Exit For
      End If
    Next z
  Next y
  
  'Loop through lines on HFA and Oracle side to mark the hits and non-hits
  'Since lines are in order, if a line ID is only half valid, it will simply turn the validation or HFA sheet red
  'Oracle
  For y = 1 To Worksheets(4).range("Y3000").End(xlUp).Row
    If Worksheets(4).range("Y" & y).Interior.Color = rgbOrange Then
      Worksheets(3).range("A" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).range("B" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).range("C" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).range("D" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbAqua
      Worksheets(3).range("P" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbAqua
    ElseIf Worksheets(4).range("Y" & y).Value <> "" Then
      Worksheets(3).range("A" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("B" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("C" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("D" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbSalmon
      Worksheets(3).range("P" & Worksheets(4).range("X" & y).Value).Interior.Color = rgbSalmon
    End If
  Next y
  
  For y = 1 To Worksheets(4).range("S3000").End(xlUp).Row
    If Worksheets(4).range("S" & y).Interior.Color = rgbOrange Then
      Worksheets(1).range("A" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).range("B" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).range("C" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).range("D" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).range("E" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
      Worksheets(1).range("F" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbAqua
    ElseIf Worksheets(4).range("S" & y).Value <> "" Then
      Worksheets(1).range("A" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("B" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("C" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("D" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("E" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
      Worksheets(1).range("F" & Worksheets(4).range("R" & y).Value).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Try one final pass through for HFA parts that are Red and Oracle Parts that got Ignored "grey"
  'This will pick up any Decimal Values missed aka lineals
  '1/16th tolerance
  'HFA
  For y = 2 To Worksheets(1).range("A3000").End(xlUp).Row
    If Worksheets(1).range("E" & y).Interior.Color = rgbSalmon Then
      'ORACLE
      For z = 2 To Worksheets(3).range("A3000").End(xlUp).Row
        If Worksheets(3).range("A" & z).Interior.Color = rgbGrey _
        And InStr(1, Worksheets(3).range("C" & z).Value, Worksheets(1).range("E" & y).Value) _
        And Worksheets(3).range("P" & z).Value >= (Worksheets(1).range("F" & y).Value - 0.0625) _
        And Worksheets(3).range("P" & z).Value <= (Worksheets(1).range("F" & y).Value + 0.0625) Then
          Worksheets(1).range("A" & y).Interior.Color = rgbAqua
          Worksheets(1).range("B" & y).Interior.Color = rgbAqua
          Worksheets(1).range("C" & y).Interior.Color = rgbAqua
          Worksheets(1).range("D" & y).Interior.Color = rgbAqua
          Worksheets(1).range("E" & y).Interior.Color = rgbAqua
          Worksheets(1).range("F" & y).Interior.Color = rgbAqua
          Worksheets(3).range("A" & z).Interior.Color = rgbAqua
          Worksheets(3).range("B" & z).Interior.Color = rgbAqua
          Worksheets(3).range("C" & z).Interior.Color = rgbAqua
          Worksheets(3).range("D" & z).Interior.Color = rgbAqua
          Worksheets(3).range("P" & z).Interior.Color = rgbAqua
          Exit For
        End If
      Next z
    End If
  Next y
  
  'Highlight the Unmatched Red in Oracle
  For y = 2 To Worksheets(3).range("A3000").End(xlUp).Row
    If Worksheets(3).range("A" & y).Interior.ColorIndex = xlNone Then
      Worksheets(3).range("A" & y).Interior.Color = rgbSalmon
      Worksheets(3).range("B" & y).Interior.Color = rgbSalmon
      Worksheets(3).range("C" & y).Interior.Color = rgbSalmon
      Worksheets(3).range("D" & y).Interior.Color = rgbSalmon
      Worksheets(3).range("P" & y).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Highlight the Unmatched Red in HFA
  For y = 2 To Worksheets(1).range("A3000").End(xlUp).Row
    If Worksheets(1).range("A" & y).Interior.ColorIndex = xlNone Then
      Worksheets(1).range("A" & y).Interior.Color = rgbSalmon
      Worksheets(1).range("B" & y).Interior.Color = rgbSalmon
      Worksheets(1).range("C" & y).Interior.Color = rgbSalmon
      Worksheets(1).range("D" & y).Interior.Color = rgbSalmon
      Worksheets(1).range("E" & y).Interior.Color = rgbSalmon
      Worksheets(1).range("F" & y).Interior.Color = rgbSalmon
    End If
  Next y
  
  'Highlight Grid Descriptions in Oracle - Informational
  For j = 2 To Worksheets(3).range("A3000").End(xlUp).Row
    If InStr(Worksheets(3).range("D" & j), "GRD") <> 0 _
    Or InStr(Worksheets(3).range("D" & j), "GRID") <> 0 _
    Or InStr(Worksheets(3).range("D" & j), "MUNTIN") <> 0 Then
      Worksheets(3).range("D" & j).Interior.Color = RGB(102, 102, 204)
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
LastCol = (Worksheets(1).range("A500").End(xlUp).Row) + 1
Debug.Print "length: "; LastCol

'Worksheets("Sheet1").Range("A1:D1").Copy Worksheets("Sheet2").Range("A1:D1")
'set the table headings
Worksheets("Validation").range("BJ4").Value = "Missing from Oracle"
Worksheets(1).range("E1").Copy Worksheets("Validation").range("BJ5")
Worksheets(1).range("F1").Copy Worksheets("Validation").range("BK5")
Worksheets(1).range("J1").Copy Worksheets("Validation").range("BL5")
Worksheets(1).range("K1").Copy Worksheets("Validation").range("BM5")

'check to see if cell is red, if so then copy that cell over to the other sheet
Do Until i = LastCol
    If ws.range("E" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("F" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("J" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("K" & i).Interior.Color = RGB(250, 128, 114) Then
    If ws.range("E" & i).Interior.Color = RGB(250, 128, 114) Then
        ws.range("E" & i).Copy Worksheets("Validation").range("BJ" & cnt)
    End If
    If ws.range("F" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.range("E" & i).Copy Worksheets("Validation").range("BJ" & cnt)
        End If
        ws.range("F" & i).Copy Worksheets("Validation").range("BK" & cnt)
    End If
    If ws.range("J" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.range("E" & i).Copy Worksheets("Validation").range("BJ" & cnt)
        End If
        ws.range("J" & i).Copy Worksheets("Validation").range("BL" & cnt)
    End If
    If ws.range("K" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("E" & i).Interior.Color = RGB(250, 128, 114) Then
            ws.range("E" & i).Copy Worksheets("Validation").range("BJ" & cnt)
        End If
        ws.range("K" & i).Copy Worksheets("Validation").range("BM" & cnt)
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
LastCol = (Worksheets(3).range("A3000").End(xlUp).Row) + 1
Debug.Print "length: "; LastCol

'Worksheets("Sheet1").Range("A1:D1").Copy Worksheets("Sheet2").Range("A1:D1")
'set the table headings
Worksheets("Validation").range("BO4").Value = "Missing from HFA"
Worksheets(3).range("C1").Copy Worksheets("Validation").range("BO5")
Worksheets(3).range("D1").Copy Worksheets("Validation").range("BP5")
Worksheets(3).range("P1").Copy Worksheets("Validation").range("BQ5")
Worksheets(3).range("Q1").Copy Worksheets("Validation").range("BR5")

'check to see if cell is red, if so then copy that cell over to the other sheet
Do Until i = LastCol
    If ws.range("C" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("D" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("P" & i).Interior.Color = RGB(250, 128, 114) Or ws.range("Q" & i).Interior.Color = RGB(250, 128, 114) Then
    If ws.range("C" & i).Interior.Color = RGB(250, 128, 114) Then
        'ws.Range("C" & i).Copy Range("BO" & cnt)
        ws.range("C" & i).Copy Worksheets("Validation").range("BO" & cnt)
        'ws.Range("C" & i).Value = Range("BO" & cnt).Value
    End If
    If ws.range("D" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.range("C" & i).Copy Worksheets("Validation").range("BO" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("D" & i).Copy Range("BP" & cnt)
         ws.range("D" & i).Copy Worksheets("Validation").range("BP" & cnt)
        'ws.Range("D" & i).Value = Range("BP" & cnt).Value
    End If
    If ws.range("P" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.range("C" & i).Copy Worksheets("Validation").range("BO" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("p" & i).Copy Range("BQ" & cnt)
        ws.range("P" & i).Copy Worksheets("Validation").range("BQ" & cnt)
        'ws.Range("P" & i).Value = Range("BQ" & cnt).Value
    End If
    If ws.range("Q" & i).Interior.Color = RGB(250, 128, 114) Then
        If Not ws.range("C" & i).Interior.Color = RGB(250, 128, 114) Then
            'ws.Range("C" & i).Copy Range("BO" & cnt)
            ws.range("C" & i).Copy Worksheets("Validation").range("BR" & cnt)
            'ws.Range("C" & i).Value = Range("BO" & cnt).Value
        End If
        'ws.Range("Q" & i).Copy Range("BR" & cnt)
        ws.range("Q" & i).Copy Worksheets("Validation").range("BR" & cnt)
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
Worksheets(3).Columns("BT:CC").Columns.AutoFit
End Sub


'=================================================================================================================================================================================================
'Check to indicate whether BOM's are a match or not
'=================================================================================================================================================================================================
Sub DoBomsMatch()

Dim item As Variant

For Each item In Worksheets(1).range("A2:R3000").Cells
    If item.Interior.Color = RGB(250, 128, 114) Then
        Worksheets("Validation").range("BJ1").Value = "ERRORS in HFA BOM"
        Worksheets("Validation").range("BJ1").Interior.Color = RGB(255, 0, 0)
        Exit For
    Else
        Worksheets("Validation").range("BJ1").Value = "NO Errors in HFA BOM"
        Worksheets("Validation").range("BJ1").Interior.Color = RGB(0, 255, 0)
    End If
Next

For Each item In Worksheets(3).range("A2:BG3000").Cells
    If item.Interior.Color = RGB(250, 128, 114) Then
        Worksheets("Validation").range("BO1").Value = "ERRORS in Oracle BOM"
        Worksheets("Validation").range("BO1").Interior.Color = RGB(255, 0, 0)
        Exit For
    Else
        Worksheets("Validation").range("BO1").Value = "NO Errors in Oracle BOM"
        Worksheets("Validation").range("BO1").Interior.Color = RGB(0, 255, 0)
    End If
Next

End Sub


'=================================================================================================================================================================================================
'Check cut angles on GT parts
'=================================================================================================================================================================================================
Sub YellowBananas()

Dim ws As Worksheet
Set ws = Sheets(1)

Dim lngLastRow As Long
Dim lngLastColumn As Long
Dim c As Long
Dim r As Long
Dim iw As Long
Dim ih As Long
Dim ipn As Long
iw = 0
ih = 0
ipn = 0

Dim arrWidth() As Double
ReDim Preserve arrWidth(iw)

Dim arrHeight() As Double
ReDim Preserve arrHeight(ih)

Dim arrPartNumber() As Variant
ReDim Preserve arrPartNumber(ipn)

Dim wsHfa As Worksheet
Set wsHfa = Sheets(1)

lngLastRow = ws.Cells(Rows.count, "A").End(xlUp).row
lngLastColumn = ws.Cells(1, Columns.count).End(xlToLeft).Column
Debug.Print "rows: " & lngLastRow

'hfa get part number
For c = 2 To lngLastColumn
    If ws.Cells(1, c).Value = "Part Number" Then
    Debug.Print "---" & ws.Cells(1, c).Value
        For r = 2 To lngLastRow
           If ws.Cells(r, c).Value Like "GA*" Or ws.Cells(r, c).Value Like "GT*" Then
                If ws.Cells(r, 10).Interior.Color <> rgbGrey Then
                    Debug.Print ws.Cells(r, 10).Value&; ":" & ws.Cells(r, c).Value
                    ReDim Preserve arrPartNumber(0 To ipn)
                    arrPartNumber(ipn) = ws.Cells(r, c).Value
                   ipn = ipn + 1

                End If
            End If
        Next r
    End If
Next c


'hfa, get unit width
Dim colHeaderWidth As String
colHeaderWidth = "Unit Width"
Dim colPosWidth As Integer
colPosWidth = 5
Call GetUnitDimensionsHfa(arrWidth, lngLastRow, lngLastColumn, iw, colHeaderWidth, colPosWidth, wsHfa)


'hfa, get unit height
Dim colHeaderHeight As String
colHeaderHeight = "Unit Hight"
Dim colPosHeight As Integer
colPosHeight = 5
Call GetUnitDimensionsHfa(arrHeight, lngLastRow, lngLastColumn, ih, colHeaderHeight, colPosHeight, wsHfa)


'print to see if getting right values
For Each item In arrWidth
    Debug.Print ("printing item width: " & item)
Next
For Each item In arrHeight
    Debug.Print ("printing item height: " & item)
Next
For Each item In arrPartNumber
    Debug.Print ("printing item pn: " & item)
Next


'going though oracle to check if the cut lengths are the same or not
Dim ws2 As Worksheet
Set ws2 = Sheets(3)

Dim temp As Double
Dim count As Integer
Dim descArr() As String
Dim oracleLastRow As Long
Dim oracleLastColumn As Long
Dim i As Integer
Dim internalCountWidth As Integer
internalCountWidth = 0
Dim internalCountHeight As Integer
internalCountHeight = 0
Dim internalCountIgu As Integer
internalCountIgu = 0
'Dim dict As New Scripting.Dictionary
'Dim dict  As Collection
'Set dict = New Collection

Dim strWidth() As Double
ReDim Preserve strWidth(internalCountWidth)
Dim strHeight() As Double
ReDim Preserve strHeight(internalCountHeight)
Dim strIgu() As Variant
ReDim Preserve strIgu(internalCountIgu)


oracleLastRow = ws2.Cells(Rows.count, "A").End(xlUp).row
oracleLastColumn = ws.Cells(1, Columns.count).End(xlToLeft).Column

For c = 2 To oracleLastColumn
    If ws2.Cells(1, c).Value = "Description" Then
    For r = 2 To oracleLastRow
        If InStr(1, ws2.Cells(r, 3), "IGU") Then
            count = 0
            Debug.Print ws2.Cells(r, 3).Value & " $$$$$ " & ws2.Cells(r, c).Value
            ReDim Preserve strIgu(0 To internalCountIgu)
            strIgu(internalCountIgu) = ws2.Cells(r, 3).Value
            internalCountIgu = internalCountIgu + 1
            descArr() = Split(ws2.Cells(r, c).Value)
            For i = LBound(descArr) To UBound(descArr)
                If Val(descArr(i)) <> 0 And count < 2 Then
                'dict.Add (ws2.Cells(r, 3).Value), (strArr(0) = descArr(i))
                    If count Mod 2 = 0 Then
                        temp = CDbl(descArr(i))
                        Debug.Print temp
                        ReDim Preserve strWidth(0 To internalCountWidth)
                        strWidth(internalCountWidth) = temp
                        internalCountWidth = internalCountWidth + 1
                    Else
                        temp = CDbl(descArr(i))
                        Debug.Print temp
                        ReDim Preserve strHeight(0 To internalCountHeight)
                        strHeight(internalCountHeight) = temp
                        internalCountHeight = internalCountHeight + 1
                    End If
                    count = count + 1
                End If
            Next i
        End If
    Next r
    End If

Next c



For Each item In strWidth
    Debug.Print ("Width arr: " & item)
Next
For Each item In strHeight
    Debug.Print ("Height arr: " & item)
Next
For Each item In strIgu
    Debug.Print ("Igu arr: " & item)
Next


'print to validiton sheet the IGU part numbers with dimensions - create variable range
Dim copyRangeStrIgu As String
StartRow = 1
LastRow = UBound(strIgu)
LastRow = LastRow + 1
copyRangeStrIgu = "BU" & StartRow & ":" & "BU" & LastRow

Dim copyRangeStrWidth As String
StartRow1 = 1
LastRow1 = UBound(strWidth)
LastRow1 = LastRow1 + 1
copyRangeStrWidth = "BV" & StartRow1 & ":" & "BV" & LastRow1

Dim copyRangeStrHeight As String
StartRow2 = 1
LastRow2 = UBound(strHeight)
LastRow2 = LastRow2 + 1
copyRangeStrHeight = "BW" & StartRow2 & ":" & "BW" & LastRow2

'"BU6:BU11"
Sheets(3).range(copyRangeStrIgu).Value = Application.Transpose(strIgu)
'Sheets(3).range("BU2").Resize((UBound(strIgu) - LBound(strIgu)) + 1, 1).Value = Application.Transpose(strIgu)
Sheets(3).range(copyRangeStrWidth).Value = Application.Transpose(strWidth)
'Sheets(3).range("BV2").Resize((UBound(strWidth) - LBound(strWidth)) + 1, 1).Value = Application.Transpose(strWidth)
Sheets(3).range(copyRangeStrHeight).Value = Application.Transpose(strHeight)
'Sheets(3).range("BW2").Resize((UBound(strHeight) - LBound(strHeight)) + 1, 1).Value = Application.Transpose(strHeight)

Dim copyRangePartNumber As String
StartRow3 = 1
LastRow3 = UBound(strIgu)
LastRow3 = LastRow3 + 1
copyRangePartNumber = "CA" & StartRow3 & ":" & "CA" & LastRow3

Dim copyRangeWidth As String
StartRow4 = 1
LastRow4 = UBound(strIgu)
LastRow4 = LastRow4 + 1
copyRangeWidth = "CB" & StartRow4 & ":" & "CB" & LastRow4

Dim copyRangeHeight As String
StartRow5 = 1
LastRow5 = UBound(strIgu)
LastRow5 = LastRow5 + 1
copyRangeHeight = "CC" & StartRow5 & ":" & "CC" & LastRow5

'Sheets(3).range(copyRangePartNumber).Value = Application.Transpose(arrPartNumber)
Sheets(3).range("CA2").Resize((UBound(arrPartNumber) - LBound(arrPartNumber)) + 1, 1).Value = Application.Transpose(arrPartNumber)
'Sheets(3).range(copyRangeWidth).Value = Application.Transpose(arrWidth)
Sheets(3).range("CB2").Resize((UBound(arrWidth) - LBound(arrWidth)) + 1, 1).Value = Application.Transpose(arrWidth)
'Sheets(3).range(copyRangeHeight).Value = Application.Transpose(arrHeight)
Sheets(3).range("CC2").Resize((UBound(arrHeight) - LBound(arrHeight)) + 1, 1).Value = Application.Transpose(arrHeight)

'print values to validation sheeet
Dim s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String, s7 As String, s8 As String, s9 As String, s10 As String
Dim r1 As Integer, r2 As Integer
r1 = LastRow + 1
r2 = LastRow3 + 1
s1 = "BT" & r1
s2 = "BU" & r1
s3 = "BV" & r1
s4 = "BW" & r1
s5 = "BX" & r1
s6 = "BY" & r1
s7 = "BZ" & r2
s8 = "CA" & r2
s9 = "CB" & r2
s10 = "CC" & r2
Worksheets("Validation").range(s1).Value = "Oracle" '+ vbNewLine + "Values for IGU parts"
Worksheets("Validation").range(s1).Interior.Color = rgbYellow
Worksheets("Validation").range(s2).Value = "IGU Part Number"
Worksheets("Validation").range(s2).Interior.Color = rgbYellow
Worksheets("Validation").range(s3).Value = "IGU Width"
Worksheets("Validation").range(s3).Interior.Color = rgbYellow
Worksheets("Validation").range(s4).Value = "IGU Height"
Worksheets("Validation").range(s4).Interior.Color = rgbYellow
'Worksheets("Validation").range(s5).Value = "Matching Dimensions"     'green for yes, red for no
'Worksheets("Validation").range(s6).Value = "Off by (tolerance max up to 1/16 or 0.0625)"

Worksheets("Validation").range("BZ1").Value = "HFA" '+ vbNewLine + "Values for glass parts"
Worksheets("Validation").range("BZ1").Interior.Color = rgbYellow
Worksheets("Validation").range("CA1").Value = "Part Number"
Worksheets("Validation").range("CA1").Interior.Color = rgbYellow
Worksheets("Validation").range("CB1").Value = "Width"
Worksheets("Validation").range("CB1").Interior.Color = rgbYellow
Worksheets("Validation").range("CC1").Value = "Height"
Worksheets("Validation").range("CC1").Interior.Color = rgbYellow


'compare oracle values with hfa values to determine whether the dimensions match or not
Dim iterex As Integer
Dim iterin As Integer
Dim iguWidth As Double
Dim iguHeight As Double
Dim isIguWidth As Boolean
isIguWidth = False
Dim isIguHeight As Boolean
isIguHeight = False
Dim correctCount As Long
correctCount = 0
Dim incorrectCount As Long
correctHCount = 0

Dim correctWidthArr() As Double
ReDim Preserve correctWidthArr(correctCount)
Dim correctHeightArr() As Double
ReDim Preserve correctHeightArr(correctCount)

'igu width to compare with hfa width
For iterex = LBound(strWidth) To UBound(strWidth)
    iguWidth = strWidth(iterex)
    Debug.Print "temp var for iguwidth: " & iguWidth
    For iterin = LBound(arrWidth) To UBound(arrWidth)
        'If iguWidth = arrWidth(iterin) Then
         '   isIguWidth = True
          '  Debug.Print "exact match: " & isIguWidth
            'Worksheets(3).range("BV" & iterex).Interior.Color = rgbGreen
           ' Exit For
        If Math.Abs(iguWidth - arrWidth(iterin)) <= 0.0625 Then
            isIguWidth = True
            Debug.Print "approx match within 1/16: " & isIguWidth
            ReDim Preserve correctWidthArr(0 To correctCount)
            correctWidthArr(correctCount) = iguWidth
            correctCount = correctCount + 1
            'Worksheets(3).range("BV" & iterex).Interior.Color = rgbGreen
            Exit For
        Else
            isIguWidth = False
            Debug.Print "not a match: " & isIguWidth
            'If IsInArray(incorrectWidthArr, iguWidth) Then
            'ReDim Preserve incorrectWidthArr(0 To incorrectCount)
            'incorrectWidthArr(incorrectCount) = iguWidth
            'incorrectCount = incorrectCount + 1
            'End If
        End If
    Next iterin
Next iterex

'igu height to compare with hfa height
For iterex = LBound(strHeight) To UBound(strHeight)
    iguHeight = strHeight(iterex)
    Debug.Print "temp var for iguheight: " & iguHeight
    For iterin = LBound(arrHeight) To UBound(arrHeight)
        If Math.Abs(iguHeight - arrHeight(iterin)) <= 0.0625 Then
            isIguHeight = True
            Debug.Print "approx match within 1/16: " & isIguHeight
            ReDim Preserve correctHeightArr(0 To correctHCount)
            correctHeightArr(correctHCount) = iguHeight
            correctHCount = correctHCount + 1
            Exit For
        Else
            isIguHeight = False
            Debug.Print "not a match: " & isIguHeight
        End If
    Next iterin
Next iterex

For Each item In correctWidthArr
Debug.Print "correct item width: " & item
Next

For Each item In correctHeightArr
Debug.Print "correct item height: " & item
Next

'go through oracle igu values and whichever values are correct, color green otherwise color red for incorrect
'check igu widths
Dim rangeWidth As range
Set rangeWidth = Sheets("Validation").range("BV:BV")
Dim likeStr1 As String
likeStr1 = "IGU Width*"
Call CompareValuesAndColorCell(rangeWidth, likeStr1, correctWidthArr)

'check igu height
Dim rangeHeight As range
Set rangeHeight = Sheets("Validation").range("BW:BW")
Dim likeStr2 As String
likeStr2 = "IGU Height*"
Call CompareValuesAndColorCell(rangeHeight, likeStr2, correctHeightArr)



End Sub

'=================================================================================================================================================================================================
'Search through array to see if correct value is present
'=================================================================================================================================================================================================
Function IsInArray(arr As Variant, item As Double) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = item Then
        IsInArray = True
        Exit Function
        End If
    Next i
    IsInArray = False
    
End Function

'=================================================================================================================================================================================================
'Compare igu width and height and color cells to indicate if correct or incorrect
'=================================================================================================================================================================================================
Function CompareValuesAndColorCell(rg As range, likeStr As String, correctArr As Variant)
For Each w In rg
    If w.Value Like likeStr Then
        Exit For
    End If
    Debug.Print w.Value
    If IsInArray(correctArr, w.Value) Then
        w.Interior.Color = rgbGreen
    Else
        w.Interior.Color = rgbRed
    End If
Next

End Function

'=================================================================================================================================================================================================
'Get the unit width and height of the glass units from HFA BOM
'=================================================================================================================================================================================================
Function GetUnitDimensionsHfa(arr As Variant, row As Long, col As Long, arrCount As Long, strColHeader As String, colPos As Integer, ws As Worksheet)
For c = 2 To col
    If ws.Cells(1, c).Value = strColHeader Then
    Debug.Print "---" & ws.Cells(1, c).Value
        For r = 2 To row
            If ws.Cells(r, colPos).Value Like "GA*" Or ws.Cells(r, colPos).Value Like "GT*" Then
            'ws.Cells(r, 5).Interior.Color = rgbOrange Then
                If ws.Cells(r, c).Interior.Color = rgbGreen Or ws.Cells(r, c).Interior.Color = rgbSalmon Then
                    Debug.Print ws.Cells(r, colPos).Value&; ":" & ws.Cells(r, c).Value
                    ReDim Preserve arr(0 To arrCount)
                    arr(arrCount) = ws.Cells(r, c).Value
                    arrCount = arrCount + 1
                End If
            End If
        Next r
    End If
Next c
End Function



