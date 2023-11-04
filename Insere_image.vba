' Ce script permet d'inserer une image dans un fichier Excel 
' en liant le contenu d'une cellule du fichier Excel à un nom de fichier jpg 
' stocké dans un dossier identifié 

Sub Insere_image()
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' Déclaration des variables
    Dim xFolder As String
    Dim xImageFile As String
    Dim xType As String
    Dim Nf As String
    Dim R As Range

    ' Valeurs des variables
    xFolder = "dossierimages/"
    xType = ".jpg"
    
    ' Définir la plage de données dans la colonne
    Dim Column_Range As Range
    Set Column_Range = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
    
    ' Variable pour numéro d'image incrémental
    Dim numImage As Integer
    numImage = 1
    
    ' Supprimer les images existantes
    Dim shape As shape
    For Each shape In ActiveSheet.Shapes
        shape.Delete
    Next shape
    
    ' Boucle à travers toutes les cellules de la plage Column_Range
    For Each R In Column_Range
        If Not IsEmpty(R) Then ' Vérifier si la cellule n'est pas vide
            xImageFile = R.Value
            Nf = xFolder & xImageFile & xType
            
            ' Vérifier si le fichier image existe
            If Dir(Nf) = "" Then
                ' MsgBox "Le fichier image spécifié n'existe pas."
            Else
                ' Référence à la cellule de destination dans la colonne "B"
                Set R = ActiveSheet.Range("B" & R.Row)
                
                With ActiveSheet
                    ' Insérer et configurer la forme image
                    Dim newShape As shape
                    Set newShape = .Shapes.AddPicture(Filename:=Nf, LinkToFile:=msoTrue, _
                       SaveWithDocument:=msoTrue, Left:=R.Left + 5, Top:=R.Top + 5, Width:=-1, Height:=-1)
                    newShape.Name = xImageFile & "_" & Format(numImage, "0000")
                    R.Value = Format(numImage, "0000")
                    newShape.LockAspectRatio = msoTrue
                    newShape.Height = 100
                    newShape.Width = 100
                    
                    ' Redimensionner la cellule de destination
                    R.RowHeight = newShape.Height + 10
                    R.ColumnWidth = 20
                    
                    ' Déplacer l'image avec les cellules
                    newShape.Placement = xlMoveAndSize
                    
                    ' Centrer le texte au milieu de la cellule
                    R.HorizontalAlignment = xlCenter
                    R.VerticalAlignment = xlCenter
                End With
            End If
            
            ' Incrémenter le numéro d'image
            numImage = numImage + 1
        End If
    Next R
        
    ExitSub:
        Application.ScreenUpdating = True
        Exit Sub
        
    ErrorHandler:
        ' Gérer l'erreur
        MsgBox "Une erreur s'est produite : " & Err.Description
        Application.ScreenUpdating = True
End Sub
