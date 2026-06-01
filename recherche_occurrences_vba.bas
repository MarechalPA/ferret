'===========================================================================
' PROGRAMME VBA : Recherche d'occurrences dans Excel
' Description : Recherche un terme dans un fichier Excel et crée un nouveau
'              fichier avec un onglet "occurence" contenant des liens fonctionnels
' Auteur : Assistant IA
' Date : 2024
'===========================================================================

Option Explicit

'===========================================================================
' PROCÉDURE PRINCIPALE
'===========================================================================
Sub RechercherOccurrencesEtCreerOngletOccurence()
    Dim wsSource As Worksheet
    Dim wbSource As Workbook
    Dim wbNew As Workbook
    Dim wsOccurence As Worksheet
    Dim searchTerm As String
    Dim filePath As String
    Dim newFileName As String
    Dim fileDialog As FileDialog
    Dim searchAllSheets As Boolean
    Dim resultCount As Long
    Dim startTime As Double
    
    ' Initialisation
    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    '===========================================================================
    ' SÉLECTION DU FICHIER
    '===========================================================================
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Sélectionnez un fichier Excel depuis D:\Documents\epenspamar\Desktop"
        .InitialFileName = "D:\Documents\epenspamar\Desktop\"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "Aucun fichier sélectionné. Opération annulée.", vbExclamation, "Annulation"
            Exit Sub
        End If
    End With
    
    '===========================================================================
    ' OUVERTURE DU FICHIER SOURCE
    '===========================================================================
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    
    '===========================================================================
    ' SAISIE DU TERME DE RECHERCHE
    '===========================================================================
    searchTerm = InputBox("Entrez le terme à rechercher :", "Recherche d'occurrences", "")
    
    If searchTerm = "" Then
        MsgBox "Aucun terme de recherche spécifié. Opération annulée.", vbExclamation, "Annulation"
        wbSource.Close SaveChanges:=False
        Exit Sub
    End If
    
    '===========================================================================
    ' OPTION : RECHERCHER DANS TOUS LES ONGLETS ?
    '===========================================================================
    searchAllSheets = (MsgBox("Voulez-vous rechercher dans TOUS les onglets ?", _
                            vbQuestion + vbYesNo, "Recherche") = vbYes)
    
    '===========================================================================
    ' CRÉATION DU NOUVEAU CLASSEUR
    '===========================================================================
    Set wbNew = Workbooks.Add
    
    ' Supprimer les feuilles par défaut
    Application.DisplayAlerts = False
    For Each wsSource In wbNew.Worksheets
        wsSource.Delete
    Next wsSource
    Application.DisplayAlerts = True
    
    ' Créer l'onglet "occurence"
    Set wsOccurence = wbNew.Worksheets.Add(After:=wbNew.Worksheets(wbNew.Worksheets.Count))
    wsOccurence.Name = "occurence"
    
    '===========================================================================
    ' RECHERCHE DES OCCURRENCES
    '===========================================================================
    resultCount = 0
    
    ' En-têtes
    With wsOccurence
        .Range("A1").Value = "Onglet"
        .Range("B1").Value = "Cellule"
        .Range("C1").Value = "Contenu"
        .Range("D1").Value = "Lien"
        
        ' Style des en-têtes
        With .Range("A1:D1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(79, 129, 189) ' Bleu
            .HorizontalAlignment = xlCenter
        End With
        
        ' Largeur des colonnes
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 50
        .Columns("D").ColumnWidth = 15
    End With
    
    ' Recherche dans les onglets
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim rowIndex As Long
    Dim sheetName As String
    Dim cellAddress As String
    Dim cellValue As String
    Dim hyperlinkFormula As String
    Dim originalFileName As String
    
    originalFileName = wbSource.Name
    rowIndex = 2 ' Commencer à la ligne 2
    
    If searchAllSheets Then
        ' Rechercher dans tous les onglets
        For Each ws In wbSource.Worksheets
            sheetName = ws.Name
            Set rng = ws.UsedRange
            
            For Each cell In rng
                If Not IsEmpty(cell.Value) Then
                    cellValue = CStr(cell.Value)
                    If InStr(1, LCase(cellValue), LCase(searchTerm), vbTextCompare) > 0 Then
                        cellAddress = cell.Address(False, False)
                        
                        ' Écrire les informations
                        wsOccurence.Cells(rowIndex, 1).Value = sheetName
                        wsOccurence.Cells(rowIndex, 2).Value = cellAddress
                        wsOccurence.Cells(rowIndex, 3).Value = Left(cellValue, 100) & _
                            IIf(Len(cellValue) > 100, "...", "")
                        
                        ' Créer le lien hypertexte
                        hyperlinkFormula = "=HYPERLINK(""[" & originalFileName & "]" & sheetName & "!" & cellAddress & "", ""Voir"")"
                        wsOccurence.Cells(rowIndex, 4).Formula = hyperlinkFormula
                        
                        ' Style de la cellule de lien
                        With wsOccurence.Cells(rowIndex, 4)
                            .Font.Color = RGB(0, 0, 255) ' Bleu
                            .Font.Underline = xlUnderlineStyleSingle
                        End With
                        
                        rowIndex = rowIndex + 1
                        resultCount = resultCount + 1
                    End If
                End If
            Next cell
        Next ws
    Else
        ' Rechercher seulement dans le premier onglet
        Set ws = wbSource.Worksheets(1)
        sheetName = ws.Name
        Set rng = ws.UsedRange
        
        For Each cell In rng
            If Not IsEmpty(cell.Value) Then
                cellValue = CStr(cell.Value)
                If InStr(1, LCase(cellValue), LCase(searchTerm), vbTextCompare) > 0 Then
                    cellAddress = cell.Address(False, False)
                    
                    ' Écrire les informations
                    wsOccurence.Cells(rowIndex, 1).Value = sheetName
                    wsOccurence.Cells(rowIndex, 2).Value = cellAddress
                    wsOccurence.Cells(rowIndex, 3).Value = Left(cellValue, 100) & _
                        IIf(Len(cellValue) > 100, "...", "")
                    
                    ' Créer le lien hypertexte
                    hyperlinkFormula = "=HYPERLINK(""[" & originalFileName & "]" & sheetName & "!" & cellAddress & "", ""Voir"")"
                    wsOccurence.Cells(rowIndex, 4).Formula = hyperlinkFormula
                    
                    ' Style de la cellule de lien
                    With wsOccurence.Cells(rowIndex, 4)
                        .Font.Color = RGB(0, 0, 255) ' Bleu
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                    
                    rowIndex = rowIndex + 1
                    resultCount = resultCount + 1
                End If
            End If
        Next cell
    End If
    
    '===========================================================================
    ' SAUVEGARDE DU NOUVEAU FICHIER
    '===========================================================================
    ' Créer le nom du nouveau fichier
    newFileName = "D:\Documents\epenspamar\Desktop\" & _
        Replace(originalFileName, ".xlsx", "") & _
        Replace(originalFileName, ".xls", "") & _
        Replace(originalFileName, ".xlsm", "") & "_occurence.xlsx"
    
    ' Sauvegarder le nouveau fichier
    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    '===========================================================================
    ' NETTOYAGE
    '===========================================================================
    wbSource.Close SaveChanges:=False
    wbNew.Close SaveChanges:=True
    
    ' Restaurer les paramètres
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Message de confirmation
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    MsgBox "Recherche terminée !" & vbCrLf & vbCrLf & _
           "Occurrences trouvées : " & resultCount & vbCrLf & _
           "Fichier créé : " & newFileName & vbCrLf & vbCrLf & _
           "Temps d'exécution : " & Format(elapsedTime, "0.00") & " secondes", _
           vbInformation, "Succès"
    
    Exit Sub
    
ErrorHandler:
    ' Restaurer les paramètres en cas d'erreur
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    If Not wbSource Is Nothing Then
        wbSource.Close SaveChanges:=False
    End If
    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If
    
    MsgBox "Une erreur est survenue : " & Err.Description, vbCritical, "Erreur"
End Sub

'===========================================================================
' PROCÉDURE POUR RECHERCHER DANS UN FICHIER SPÉCIFIQUE
' (Alternative avec chemin fixe)
'===========================================================================
Sub RechercherOccurrencesFichierSpecifique()
    Dim filePath As String
    Dim searchTerm As String
    
    ' Chemin fixe vers le répertoire
    filePath = "D:\Documents\epenspamar\Desktop\EdT 2026-27 FISE 2A V2.xlsx"
    
    ' Vérifier si le fichier existe
    If Dir(filePath) = "" Then
        MsgBox "Le fichier n'existe pas : " & filePath, vbExclamation, "Fichier introuvable"
        Exit Sub
    End If
    
    ' Demander le terme de recherche
    searchTerm = InputBox("Entrez le terme à rechercher :", "Recherche d'occurrences", "")
    
    If searchTerm = "" Then
        MsgBox "Aucun terme de recherche spécifié.", vbExclamation, "Annulation"
        Exit Sub
    End If
    
    ' Appeler la procédure principale avec le fichier spécifique
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    
    ' Créer un nouveau classeur pour les résultats
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add
    
    ' Supprimer les feuilles par défaut
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In wbNew.Worksheets
        ws.Delete
    Next ws
    Application.DisplayAlerts = True
    
    ' Créer l'onglet "occurence"
    Dim wsOccurence As Worksheet
    Set wsOccurence = wbNew.Worksheets.Add
    wsOccurence.Name = "occurence"
    
    ' Rechercher dans tous les onglets
    Call RechercherDansOnglets(wbSource, wbNew, wsOccurence, searchTerm)
    
    ' Sauvegarder le nouveau fichier
    Dim newFileName As String
    newFileName = "D:\Documents\epenspamar\Desktop\EdT_2026-27_FISE_2A_V2_occurence.xlsx"
    
    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    
    ' Fermer les fichiers
    wbSource.Close SaveChanges:=False
    wbNew.Close SaveChanges:=True
    
    MsgBox "Recherche terminée !" & vbCrLf & vbCrLf & _
           "Fichier créé : " & newFileName, _
           vbInformation, "Succès"
End Sub

'===========================================================================
' FONCTION DE RECHERCHE DANS LES ONGLETS
'===========================================================================
Private Sub RechercherDansOnglets(wbSource As Workbook, wbNew As Workbook, _
                                wsOccurence As Worksheet, searchTerm As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim rowIndex As Long
    Dim sheetName As String
    Dim cellAddress As String
    Dim cellValue As String
    Dim hyperlinkFormula As String
    Dim originalFileName As String
    Dim resultCount As Long
    
    originalFileName = wbSource.Name
    rowIndex = 2
    resultCount = 0
    
    ' En-têtes
    With wsOccurence
        .Range("A1").Value = "Onglet"
        .Range("B1").Value = "Cellule"
        .Range("C1").Value = "Contenu"
        .Range("D1").Value = "Lien"
        
        ' Style des en-têtes
        With .Range("A1:D1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(79, 129, 189)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Largeur des colonnes
        .Columns("A").ColumnWidth = 15
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 50
        .Columns("D").ColumnWidth = 15
    End With
    
    ' Rechercher dans tous les onglets
    For Each ws In wbSource.Worksheets
        sheetName = ws.Name
        Set rng = ws.UsedRange
        
        For Each cell In rng
            If Not IsEmpty(cell.Value) Then
                cellValue = CStr(cell.Value)
                If InStr(1, LCase(cellValue), LCase(searchTerm), vbTextCompare) > 0 Then
                    cellAddress = cell.Address(False, False)
                    
                    ' Écrire les informations
                    wsOccurence.Cells(rowIndex, 1).Value = sheetName
                    wsOccurence.Cells(rowIndex, 2).Value = cellAddress
                    wsOccurence.Cells(rowIndex, 3).Value = Left(cellValue, 100) & _
                        IIf(Len(cellValue) > 100, "...", "")
                    
                    ' Créer le lien hypertexte
                    hyperlinkFormula = "=HYPERLINK(""[" & originalFileName & "]" & sheetName & "!" & cellAddress & "", ""Voir"")"
                    wsOccurence.Cells(rowIndex, 4).Formula = hyperlinkFormula
                    
                    ' Style de la cellule de lien
                    With wsOccurence.Cells(rowIndex, 4)
                        .Font.Color = RGB(0, 0, 255)
                        .Font.Underline = xlUnderlineStyleSingle
                    End With
                    
                    rowIndex = rowIndex + 1
                    resultCount = resultCount + 1
                End If
            End If
        Next cell
    Next ws
    
    ' Message si aucune occurrence trouvée
    If resultCount = 0 Then
        wsOccurence.Cells(2, 1).Value = "Aucune occurrence trouvée pour : " & searchTerm
        wsOccurence.Range("A2:D2").Merge
        wsOccurence.Cells(2, 1).HorizontalAlignment = xlCenter
        wsOccurence.Cells(2, 1).Font.Italic = True
        wsOccurence.Cells(2, 1).Font.Color = RGB(128, 128, 128)
    End If
End Sub

'===========================================================================
' INSTRUCTIONS D'INSTALLATION ET D'UTILISATION
'===========================================================================
' 
' POUR UTILISER CE PROGRAMME :
' 
' 1. Ouvrez Excel
' 2. Appuyez sur ALT + F11 pour ouvrir l'éditeur VBA
' 3. Dans l'éditeur VBA :
'    - Cliquez sur Insertion > Module
'    - Copiez tout le code ci-dessus
'    - Collez-le dans le module
' 4. Fermez l'éditeur VBA
' 5. Pour exécuter :
'    - Appuyez sur ALT + F8
'    - Sélectionnez "RechercherOccurrencesEtCreerOngletOccurence"
'    - Cliquez sur "Exécuter"
' 
' OU :
'    - Créez un bouton dans une feuille Excel
'    - Assignez-lui la macro "RechercherOccurrencesEtCreerOngletOccurence"
' 
' FONCTIONNALITÉS :
' - Recherche insensible à la casse
' - Création d'un NOUVEAU fichier avec onglet "occurence"
' - Liens fonctionnels vers les cellules originales
' - Style appliqué aux en-têtes et liens
' - Gestion des erreurs
' 
'===========================================================================
