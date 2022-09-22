Attribute VB_Name = "EASAREADOUT"
Function getRGB1(FCell As Range) As String ' returns a string with the color code of the cell


'Nota Bene : "0072EC2" = Regulations ; "212E63"= CS


    Dim xColor As String
    xColor = CStr(FCell.Interior.Color)
    xColor = Right("000000" & Hex(xColor), 6)
    getRGB1 = Right(xColor, 2) & Mid(xColor, 3, 2) & Left(xColor, 2)
    
End Function



Sub getlistofreg()




Set aim = ThisWorkbook.Worksheets("Reg")
Set src1 = ThisWorkbook.Worksheets("Table 1")


regn = 1



For k = 4 To 580






' InStr(1, src1.Cells(k, 1), "Regulation (EU)") And











'Or getRGB1(Cells(k, 1)) = "212E63"


    If getRGB1(src1.Cells(k, 1)) = "007EC2" Then
    
  'src1.Cells(k, 1)= Reg ID + Name of the reg
  
    
    
    '----------------------------------           ' title of regulation------------------------------------
    
            aim.Cells(regn, 1) = Split(CStr(src1.Cells(k, 1)), " ")(0)
            
            If InStr(1, CStr(src1.Cells(k + 1, 1)), Chr(10)) Then
            
              aim.Cells(regn, 4) = Split(CStr(src1.Cells(k, 1)), " ")(1) & " - " & Split(CStr(src1.Cells(k + 1, 1)), Chr(10))(0)
              
            Else
            
            
            
            If src1.Cells(k, 1) <> "Regulation" Then
            
            
            aim.Cells(regn, 4) = Split(CStr(src1.Cells(k, 1)), " ")(1) & " - " & CStr(src1.Cells(k + 1, 1))
            
            End If
            
            
            End If
            
            
            ' title of regulation
       '------------------------------------- content -----------------------------
       
       
       contentreg = Split(CStr(src1.Cells(k + 1, 1)), Chr(10))
        For w = 1 To UBound(contentreg)
        
             If Left(contentreg(w), 1) <> "(" Then
        
                   aim.Cells(regn, 4) = aim.Cells(regn, 4) & Chr(10) & contentreg(w)
                    
                    
            End If
            
        
        
        
                If Left(contentreg(w), 1) = "(" Then
                
                
                      If Left(contentreg(w), 2) <> "(i" And Left(contentreg(w), 2) <> "(v" Then
                            
                            regn = regn + 1
                                aim.Cells(regn, 2) = Split(contentreg(w), " ")(0)
                                aim.Cells(regn, 4) = Replace(contentreg(w), Split(contentreg(w), " ")(0), "")
                                
                            
                            
                            
                        End If
                         If Left(contentreg(w), 2) = "(i" Or Left(contentreg(w), 2) = "(v" Then
                            
                            regn = regn + 1
                                aim.Cells(regn, 3) = Split(contentreg(w), " ")(0)
                                aim.Cells(regn, 4) = Replace(contentreg(w), Split(contentreg(w), " ")(0), "")
                                
                            
                            
                            
                        End If
                        
                        
                '--- content with paragraphs
                
                
                End If
                
                
              
                
              
              
              
              
        Next
        
            
   '---------------------------- if the next cell also contains data
   
   If (Right(src1.Cells(k + 1, 1), 1) <> "." And getRGB1(src1.Cells(k + 2, 1)) <> "007EC2") Then
   condition2 = True
   Else
   condition2 = False
   End If
   
   
    If (Left(src1.Cells(k + 2, 1), 1) = "(" And getRGB1(src1.Cells(k + 2, 1)) <> "007EC2") Then
   condition3 = True
   Else
   condition3 = False
   End If
   
   
   
   
   
    If src1.Cells(k + 1, 1) <> "" And src1.Cells(k + 2, 1) <> "" Then
    
        If Split(CStr(src1.Cells(k + 1, 1)), Chr(10))(0) = Split(CStr(src1.Cells(k + 2, 1)), Chr(10))(0) Or condition2 = True Or condition3 = True Then
         
         
         
         k = k + 1
          '------------------------------------- content -----------------------------
       
       
       contentreg = Split(CStr(src1.Cells(k + 1, 1)), Chr(10))
        For w = 1 To UBound(contentreg)
        
             If Left(contentreg(w), 1) <> "(" And InStr(contentreg(w), "SUBPART") = 0 Then
        
                   aim.Cells(regn, 4) = aim.Cells(regn, 4) & Chr(10) & contentreg(w)
                    
                    
            End If
            
        
        
        
                If Left(contentreg(w), 1) = "(" Then
                
                
                        If Left(contentreg(w), 2) <> "(i" And Left(contentreg(w), 2) <> "(v" Then
                            
                            regn = regn + 1
                                aim.Cells(regn, 2) = Split(contentreg(w), " ")(0)
                                aim.Cells(regn, 4) = Replace(contentreg(w), Split(contentreg(w), " ")(0), "")
                                
                            
                            
                            
                        End If
                         If Left(contentreg(w), 2) = "(i" Or Left(contentreg(w), 2) = "(v" Then
                            
                            regn = regn + 1
                                aim.Cells(regn, 3) = Split(contentreg(w), " ")(0)
                                aim.Cells(regn, 4) = Replace(contentreg(w), Split(contentreg(w), " ")(0), "")
                                
                            
                            
                            
                        End If
                        
                        
                '--- content with paragraphs
                
                
                End If
                
                Next
                
         
         
         
         End If
         
      End If
      
         
         
    
   
   
   
   
         
         
         
         
         
         
         
            
            
          '  aim.Cells(regn, 2) = src1.Cells(k + 1, 1)
            
            regn = regn + 1
            
            
            
            
    End If
    
Next









End Sub


Sub getlistofcsr()







Set aim = ThisWorkbook.Worksheets("CS")
Set src1 = ThisWorkbook.Worksheets("Table 1")




'-------------------------correction




























regn = 1


For k = 4 To 500


If getRGB1(src1.Cells(k, 1)) = "212E63" And InStr(1, src1.Cells(k, 1), "Part ") = 0 Then







  aim.Cells(regn, 1) = Split(CStr(src1.Cells(k, 1)), " ")(0) & " " & Split(CStr(src1.Cells(k, 1)), " ")(1)
   
   
   If InStr(1, src1.Cells(k, 1), "and") <> 0 Then
   
If Split(CStr(src1.Cells(k, 1)), " ")(2) = "and" Then


 aim.Cells(regn, 1) = Split(CStr(src1.Cells(k, 1)), " ")(0) & " " & Split(CStr(src1.Cells(k, 1)), " ")(1) & " " & Split(CStr(src1.Cells(k, 1)), " ")(2) & " " & Split(CStr(src1.Cells(k, 1)), " ")(3)
 

End If


    End If



   
    If InStr(1, aim.Cells(regn, 1), "(") Then
            aim.Cells(regn, 1) = Split(aim.Cells(regn, 1), "(")(0)
            
            
                
                    
                    
                    
                
            
   End If
   
   
   aim.Cells(regn, 5) = Replace(src1.Cells(k, 1), aim.Cells(regn, 1), "")
   
  
   
       If InStr(1, aim.Cells(regn, 5), ")") Then
                    aim.Cells(regn, 5) = Split(aim.Cells(regn, 5), ")")(UBound(Split(aim.Cells(regn, 5), ")")))
                    
                End If











    If InStr(1, Split(src1.Cells(k + 1, 1), Chr(10))(0), "ED Decision") <> 0 Or InStr(1, Split(src1.Cells(k + 1, 1), Chr(10))(0), "Regulation") <> 0 Then
    
    
    aim.Cells(regn, 5) = aim.Cells(regn, 5) & " - " & Split(src1.Cells(k + 1, 1), Chr(10))(0)
    
    End If

    



'------------------------------------ content


contentcs = Split(src1.Cells(k + 1, 1), Chr(10))

    

If UBound(contentcs) <> 0 Then

For w = 0 To UBound(contentcs)

    If InStr(1, aim.Cells(regn, 5), contentcs(w)) = 0 And InStr(1, contentcs(w), ")   ") = 0 Then
        aim.Cells(regn, 5) = aim.Cells(regn, 5) & Chr(10) & contentcs(w)
        
    
    End If
    
    
     If Left(contentcs(w), 1) = "(" And InStr(1, contentcs(w), ")   ") Then
       
        
           If IsNumeric(Replace(Left(contentcs(w), 2), "(", "")) = False And InStr(1, contentcs(w), ")   ") <> 0 Then
                 regn = regn + 1
               
                        
                        If Replace(Left(contentcs(w), 2), "(", "") = "i" Or Left(contentcs(w), 2) = "v" Then
                          aim.Cells(regn, 4) = Split(contentcs(w), ")")(0) & ")"
                    aim.Cells(regn, 5) = Replace(contentcs(w), aim.Cells(regn, 4), "")
                       Else
                        aim.Cells(regn, 2) = Split(contentcs(w), ")")(0) & ")"
                    aim.Cells(regn, 5) = Replace(contentcs(w), aim.Cells(regn, 2), "")
                    
                      End If
                    
                        
                        
           Else
           regn = regn + 1
           
          
                    aim.Cells(regn, 3) = Split(contentcs(w), ")")(0) & ")"
                    aim.Cells(regn, 5) = Replace(contentcs(w), aim.Cells(regn, 3), "")
                        
    
            End If
            
    
    End If
    
                
    
    
Next

   End If
   











           
 regn = regn + 1 ' new row for CS
 


End If








Next



End Sub


Sub copypate()


    Sheets("Table 1").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Table 12").Select
    Sheets("Table 12").Copy Before:=Sheets(1)
    Sheets("Table 12 (2)").Select
    Sheets("Table 12 (2)").Name = "Table 1"






End Sub




Sub correction()

Set src1 = ThisWorkbook.Worksheets("Table 1")


  For k = 4 To 500
  If src1.Cells(k, 1) = "Table 1" Then
src1.Cells(k, 1).EntireRow.Delete
End If
Next

For k = 2 To 500




If InStr(1, src1.Cells(k, 1), "Powered by EASA") Then


    src1.Cells(k + 1, 1).EntireRow.Delete
  src1.Cells(k, 1).EntireRow.Delete

End If



Next





For k = 2 To 500




If InStr(1, src1.Cells(k, 1), "Powered by EASA") Then


    src1.Cells(k + 1, 1).EntireRow.Delete
  src1.Cells(k, 1).EntireRow.Delete

End If


Next






For k = 2 To 500



If Right(CStr(src1.Cells(k, 1)), 1) = ":" And InStr(1, CStr(src1.Cells(k - 1, 1)), "Regulation (EU)") <> 0 Then


src1.Cells(k - 1, 1) = src1.Cells(k - 1, 1) & Chr(10) & src1.Cells(k, 1) & Chr(10) & src1.Cells(k + 1, 1)


End If


'If Right(CStr(src1.Cells(k, 1)), 1) <> "." And InStr(1, src1.Cells(k + 2, 1), ")      ") <> 0 <> "" And InStr(1, src1.Cells(k, 1), ")      ") <> 0 <> "" Then

'Debug.Print k





'---------------------------- if reg is at the end of a page


If src1.Cells(k, 1) = "" And src1.Cells(k + 1, 1) = "" Then
src1.Cells(k, 1).EntireRow.Delete
End If





If src1.Cells(k, 1) = 0 Or src1.Cells(k, 1) = "Table 1" And src1.Cells(k + 1, 1) <> "" Then
src1.Cells(k, 1).EntireRow.Delete
End If



Next


For k = 4 To 500
If Right(src1.Cells(k, 1), 1) <> "." And InStr(src1.Cells(k, 1), ")     ") <> 0 And InStr(src1.Cells(k + 1, 1), ")     ") <> 0 And InStr(1, src1.Cells(k, 1), src1.Cells(k + 1, 1)) = 0 Then

src1.Cells(k, 1) = src1.Cells(k, 1) & Chr(10) & src1.Cells(k + 1, 1)
src1.Cells(k + 1, 1).EntireRow.Delete


End If
Next



For k = 2 To 500



If Right(CStr(src1.Cells(k, 1)), 1) = "," And InStr(1, CStr(src1.Cells(k, 1)), src1.Cells(k + 1, 1)) = 0 Then


src1.Cells(k, 1) = src1.Cells(k, 1) & src1.Cells(k + 1, 1)
src1.Cells(k + 1, 1).EntireRow.Delete




End If


Next




For k = 2 To 500



If Right(CStr(src1.Cells(k, 1)), 3) = "and" And InStr(1, CStr(src1.Cells(k, 1)), src1.Cells(k + 1, 1)) = 0 Then


src1.Cells(k, 1) = src1.Cells(k, 1) & src1.Cells(k + 1, 1)
src1.Cells(k + 1, 1).EntireRow.Delete



End If


Next




For k = 2 To 500



    If src1.Cells(k, 1) = "" And src1.Cells(k + 1, 1) = "" Then
    j = 1
        Do Until src1.Cells(k, 1) <> "" Or j = 10
            src1.Cells(k, 1).EntireRow.Delete
            k = k + 1
            j = j + 1
        Loop
    End If
    
   Next

For k = 2 To 500


If InStr(1, src1.Cells(k, 1), ")   ") And InStr(1, src1.Cells(k + 1, 1), ")   ") And getRGB1(src1.Cells(k, 1)) = "FFFFFF" And getRGB1(src1.Cells(k + 1, 1)) = "FFFFFF" And InStr(1, src1.Cells(k, 1), src1.Cells(k + 1, 1)) = 0 Then

src1.Cells(k, 1) = src1.Cells(k, 1) & Chr(10) & src1.Cells(k + 1, 1)

src1.Cells(k + 1, 1).EntireRow.Delete
End If
Next

    

End Sub



Sub getlistofguid()


Set src1 = ThisWorkbook.Worksheets("Table 1")

Set aim = ThisWorkbook.Worksheets("GM")
regn = 1

For k = 1 To 500



    If getRGB1(src1.Cells(k, 1)) = "16CC7E" Then
    
        aim.Cells(regn, 1) = Split(src1.Cells(k, 1), " ")(0) & " " & Split(src1.Cells(k, 1), " ")(1)
        aim.Cells(regn, 2) = Replace(src1.Cells(k, 1), aim.Cells(regn, 1), "") & " - " & Split(src1.Cells(k + 1, 1), Chr(10))(0)
        regn = regn + 1
        Do Until getRGB1(src1.Cells(k + 1, 1)) <> "FFFFFF" Or InStr(1, src1.Cells(k + 1, 1), "FIGURE 1") <> 0
        
        aim.Cells(regn, 2) = aim.Cells(regn, 2) & Chr(10) & src1.Cells(k + 1, 1)
        
        k = k + 1
        Loop
        
        
        regn = regn + 1
    End If
    
    
Next







End Sub
