Option Explicit

Public irow As Long
Public CurrentHost As Object

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Sub CLEAR()
Set aRange = Sheets("TLO BOT").Range("A5.ZZ500000")
aRange.ClearContents
End Sub

Public Sub CBC_MASTER()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession
irow = 5
Do

  If Range("B" & irow).Value = "" Then
    Application.StatusBar = "CBC Add Complete"
    MsgBox "Add Complete"
    Exit Sub
    End If
    
If Range("A" & irow).Value = "DONE" Then
    Sleep 1
ElseIf Range("AT" & irow).Value = "" & Range("AV" & irow).Value = "" & Range("AX" & irow).Value = "" & Range("AZ" & irow).Value = "" & Range("BB" & irow).Value = "" Then
    Sleep 1
Else
    TwentySixScreen
    
    Sleep 100
    
    If Range("AT" & irow).Value = "" Then
        Sleep 1
    Else
        NumberOne
    End If
    
    Sleep 100
    
    If Range("AV" & irow).Value = "" Then
        Sleep 1
    Else
        NumberTwo
    End If
    
    Sleep 100
    
    If Range("AX" & irow).Value = "" Then
        Sleep 1
    Else
        NumberThree
    End If
    
    If Range("AZ" & irow).Value = "" Then
        Sleep 1
    Else
        NumberFour
    End If
    
    If Range("BB" & irow).Value = "" Then
        Sleep 1
    Else
        NumberFive
    End If
    
        ExitStar

End If

Range("A" & irow).Value = "DONE"
    
    irow = irow + 1
Loop

End Sub

Public Sub TwentySixScreen()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

Sleep 600
        
        If CurrentHost.GetText(0, 22, 52) = "ENTER SELECTION (.,FILE#,/,STATUS,-nnnnn,Tn,/R,HELP)" Then
        CurrentHost.Output Range("B" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output Range("B" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "26" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "26" & ChrW$(13)
        End If
  
End Sub

Public Sub NumberOne()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "CBC" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "CBC" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/11" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/11" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "335" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "335" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/14" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/14" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AT" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AT" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "//" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "//" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
        CurrentHost.Output ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ChrW$(13)
        End If
          
End Sub

Public Sub NumberTwo()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "CBC" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "CBC" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/11" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/11" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "335" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "335" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/14" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/14" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AV" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AV" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "//" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "//" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
        CurrentHost.Output ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ChrW$(13)
        End If
          
End Sub

Public Sub NumberThree()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "CBC" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "CBC" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/11" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/11" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "335" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "335" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/14" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/14" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AX" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AX" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "//" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "//" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
        CurrentHost.Output ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ChrW$(13)
        End If
          
End Sub

Public Sub NumberFour()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "CBC" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "CBC" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/11" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/11" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "335" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "335" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/14" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/14" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AZ" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("AZ" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "//" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "//" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
        CurrentHost.Output ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ChrW$(13)
        End If
          
End Sub


Public Sub NumberFive()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/A" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/A" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 47) = "DO YOU WANT TO ADD AN ADDITIONAL CONTACT? (Y,N)" Then
        CurrentHost.Output "Y" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "Y" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/1" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/1" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X,/H)" Then
        CurrentHost.Output "CBC" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "CBC" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/11" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/11" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 20) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "335" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "335" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "/14" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/14" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("BB" & irow).Value & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ThisWorkbook.Worksheets("TLO BOT").Range("BB" & irow).Value & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER (/,//,/n,X)" Then
        CurrentHost.Output "//" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "//" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 28) = "ENTER (/,W,/F,/B,SCREEN#,/n)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 19) = "OK TO FILE (CR=Y,/)" Then
        CurrentHost.Output ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output ChrW$(13)
        End If
          
End Sub

Public Sub ExitStar()
Set CurrentHost = GetObject(, "ATWin32.AccuTerm")
Set CurrentHost = CurrentHost.ActiveSession

        If CurrentHost.GetText(0, 22, 27) = "SELECTION (n,/A,/T,/F,/H,/)" Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "4" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "4" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER WHAT (nn,X)" Then
        CurrentHost.Output "16" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "16" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 14) = "ENTER WHO (nn)" Then
        CurrentHost.Output "17" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "17" & ChrW$(13)
        End If
        
        Sleep 300
        
        CurrentHost.Output "CBC LEADS ADDED" & ChrW$(13)
        Sleep 50
        CurrentHost.Output ChrW$(13)
        Sleep 50
        CurrentHost.Output ChrW$(13)
        Sleep 50
        CurrentHost.Output ChrW$(13)
        Sleep 50
        CurrentHost.Output ChrW$(13)
        
        If CurrentHost.GetText(0, 22, 48) = "ENTER SELECTION, FILE#,HELP,W,V,C,S,Dn,GC#,/,-,." Then
        CurrentHost.Output "/" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "/" & ChrW$(13)
        End If
        
        If CurrentHost.GetText(0, 22, 17) = "ENTER RESULT (nn)" Then
        CurrentHost.Output "12" & ChrW$(13)
        Else
        Sleep 200
        CurrentHost.Output "12" & ChrW$(13)
        End If
        
End Sub





