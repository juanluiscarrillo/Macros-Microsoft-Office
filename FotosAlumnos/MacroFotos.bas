Attribute VB_Name = "MacroFotos"

Sub CargarFotos()
    Dim Ruta As String
    Ruta = ThisWorkbook.Path & "\fotos\"
    
    Dim Fs
    Set Fs = CreateObject("Scripting.FileSystemObject")
    
    Dim Carpeta
    Set Carpeta = Fs.GetFolder(Ruta)
    
    ActiveSheet.Cells.Clear
    For Each Pic In ActiveSheet.Pictures
        Pic.Delete
    Next Pic
    
    Dim Contador
    Contador = 0
    ActiveSheet.Range("b4").Select
    For Each archivo In Carpeta.Files
        ActiveCell.EntireColumn.ColumnWidth = 16
        Set foto = ActiveSheet.Pictures.Insert(archivo.Path)
        
        With foto
            .ShapeRange.LockAspectRatio = msoFalse
            .ShapeRange.Height = 80
            .ShapeRange.Width = 80
        End With
                
                
        With foto
            .Top = ActiveCell.Top + 2
            .Left = ActiveCell.Left + ActiveCell.Width / 2 - .Width / 2
            ActiveCell.EntireRow.RowHeight = .Height + 4
            .Placement = xlMoveAndSize
        End With
        ActiveCell.BorderAround Color:=0, Weight:=2
        ActiveCell.Offset(1, 0).Select
        ActiveCell.RowHeight = 22
        Dim NombreAlumno() As String
        NombreAlumno = Split(archivo.Name, ".")
        NombreAlumno = Split(NombreAlumno(0), " ")
        Dim Texto As String
        Texto = ""
        Dim HayRetorno
        HayRetorno = 0
        For Each palabra In NombreAlumno
            Texto = Texto & palabra & " "
            If HayRetorno = 0 And Len(Texto) > 12 Then
                Texto = Texto & vbLf
                HayRetorno = 1
            End If
        Next palabra
        ActiveCell.Font.Name = "Arial"
        ActiveCell.Font.Size = 6
        ActiveCell = Texto
        ActiveCell.BorderAround Color:=0, Weight:=2
        ActiveCell.Offset(-1, 1).Select
        
        Contador = Contador + 1
        If Contador = 5 Then
            ActiveCell.Offset(2, -5).Select
            ActiveCell.RowHeight = 6
            ActiveCell.Offset(1, 0).Select
            Contador = 0
        End If
        
    Next archivo
    
    Range("b3").Select
    ActiveCell.Font.Name = "Arial"
    ActiveCell.Font.Size = 8
    ActiveCell.Font.Bold = True
    ActiveCell.RowHeight = 15
    ActiveCell = "GRUPO:"
    
    Range("d2").Select
    ActiveCell.Font.Name = "Arial"
    ActiveCell.Font.Size = 8
    ActiveCell.Font.Bold = True
    ActiveCell.RowHeight = 15
    ActiveCell.HorizontalAlignment = xlCenter
    ActiveCell = "LISTADO DE ALUMNOS"
    
    Range("F3").Select
    ActiveCell.Font.Name = "Arial"
    ActiveCell.Font.Size = 8
    ActiveCell.Font.Bold = True
    ActiveCell.HorizontalAlignment = xlRight
    ActiveCell = "CURSO ACADÉMICO 19-20"
    
    Range("a1").Select
    ActiveCell.ColumnWidth = 1
    
    Range("g1").Select
    ActiveCell.ColumnWidth = 1
    
    Ruta = ThisWorkbook.Path
    Range("b1").Select
    Set foto = ActiveSheet.Pictures.Insert(Ruta & "\logoies.jpg")
    
    With foto
        .ShapeRange.LockAspectRatio = msoTrue
        .ShapeRange.Height = 35
    End With
    
    With foto
        .Top = ActiveCell.Top
        .Left = ActiveCell.Left
        .Placement = xlMoveAndSize
    End With
    ActiveCell.RowHeight = foto.Height
    
    Range("f1").Select
    Set foto = ActiveSheet.Pictures.Insert(Ruta & "\logocm.jpg")
    
    With foto
        .ShapeRange.LockAspectRatio = msoTrue
        .ShapeRange.Height = 35
    End With
    With foto
        .Top = ActiveCell.Top
        .Left = ActiveCell.Left + ActiveCell.Width - .Width
        .Placement = xlMoveAndSize
    End With
    Application.ScreenUpdating = True

End Sub

