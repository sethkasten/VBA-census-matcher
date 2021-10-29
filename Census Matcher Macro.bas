Attribute VB_Name = "Module1"
Sub Matchpeople()

Dim xrow As Integer, SSN As String, FName As String, LName As String, Gender As String, NCells As Integer, xrow2 As Integer
Dim A As Integer, B As Integer, C As Integer, D As Integer, E As Integer, F As Integer, G As Integer, H As Integer, I As Integer
Dim List2 As Integer

List2 = 0
NCells = Range("K2").Value

Do While List2 < 2

    xrow = 3
    xrow2 = 3

    If List2 = 0 Then
        A = 1
        B = 2
        C = 3
        D = 4
        E = 5
        F = 6
        G = 7
        H = 8
        I = 9
    Else
        A = 5
        B = 6
        C = 7
        D = 8
        E = 1
        F = 2
        G = 3
        H = 4
        I = 10
    End If
    
    Do While xrow <= NCells
    
        xrow2 = 3
    
        SSN = Cells(xrow, A)
        FName = Cells(xrow, B)
        LName = Cells(xrow, C)
        Gender = Cells(xrow, D)
        
        Do While xrow2 <= NCells
        
            If SSN = Cells(xrow2, E) And FName = Cells(xrow2, F) And LName = Cells(xrow2, G) And Gender = Cells(xrow2, H) Then
            
                Cells(xrow, I) = "Yes"
            
            End If
            
            xrow2 = xrow2 + 1
            
        Loop
    
        xrow = xrow + 1
    
    Loop
    
    xrow = 3
    
    Do While xrow <= NCells
    
        If Cells(xrow, I) <> "Yes" And Cells(xrow, A) <> "" Then
            Cells(xrow, I) = "No"
        End If
        
        xrow = xrow + 1
    
    Loop
    
    List2 = List2 + 1
    
Loop

ActiveSheet.UsedRange

End Sub

Sub SSNFixer()

Dim xrow As Integer, NCells As Integer, ShortLen As Boolean

Columns("A:A").NumberFormat = "@"

xrow = 2

NCells = Range("B2").Value

Do While xrow <= NCells

    ShortLen = True

    Do While ShortLen = True
    
        If Len(Cells(xrow, 1)) < 4 Then
            Cells(xrow, 1) = "0" & Cells(xrow, 1)
        End If
        
        If Len(Cells(xrow, 1)) >= 4 Then
            ShortLen = False
        End If
        
    Loop

    xrow = xrow + 1

Loop

ActiveSheet.UsedRange

End Sub


