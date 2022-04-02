Private Sub CommandButton1_Click()
    
    ' TODO GENERER DEVIS / FACTURE / BON DE COMMANDE
    
    
    
    ' Je commence par insérer la date dans la cellule dédiée
    Range("C16") = Date
    
    
    ' Pour chaque cellule de la liste, je vérifie si elle n'est pas vide, si sa valeur est bien numérique et si elle est positive
    ' Si elle est vide, non numérique ou négative, je change sa couleur, je notifie l'utilisateur et j'arrête le script
    
    For Each mycell In Array(Range("C15"), Range("C26"), Range("G26"))
        If IsEmpty(mycell.Value) Then
            mycell.Interior.ColorIndex = 3
            CreateObject("WScript.Shell").PopUp "La case """ & mycell.Offset(ColumnOffset:=-1).Value & """ est vide.", 2, _
            "Cases manquantes", 0
            
            Exit Sub
        
        ElseIf Not IsNumeric(mycell.Value) Then
            CreateObject("WScript.Shell").PopUp "La cellule """ & mycell.Offset(ColumnOffset:=-1).Value & _
            """ doit être au format numérique.", 2, _
            "Mauvais format", 0
            
            Exit Sub
        
        ElseIf Not CInt(mycell.Value) > 0 Then
                CreateObject("WScript.Shell").PopUp "Les valeurs insérées doivent être supérieures à 0", 5, _
                "Quantité invalide", 48
                Exit Sub

        ElseIf Not IsEmpty(mycell.Value) Or IsNumeric(mycell.Value) Then
            mycell.Interior.ColorIndex = 2
        End If
             
    Next mycell
    
    
    ' NUMERO FACTURE
    ' J'ajoute le numéro de facture dans la case dédiée si la case est vide
    ' J'empêche l'utilisateur de changer de numéro de facture avec un panier rempli
    
    If IsEmpty(Range("J22").Value) Then
        Range("J22").Value = CInt(Range("C15").Value)
    ElseIf Not IsEmpty(Range("J22").Value) And Range("J22").Value <> Range("C15").Value Then
        CreateObject("WScript.Shell").PopUp "Avant de passer à la facture suivante, vous devez exporter le panier client puis " & _
        "le réinitialiser.", 10, _
            "Facture non terminée", 48
        
    End If
    
    
    ' NUMERO CLIENT
    ' Si la case est vide, j'insère le numéro client
    ' Si elle n'est pas vide, j'empêche l'utilisateur de changer de client si des articles sont dans le panier
    
    If IsEmpty(Range("K22").Value) Then
        Range("K22").Value = Range("C18").Value
    ElseIf Not IsEmpty(Range("K22").Value) And Range("K22").Value <> Range("C18").Value Then
        CreateObject("WScript.Shell").PopUp "Avant de modifier le numéro du client, vous devez exporter son panier puis " & _
        "le réinitialiser.", 10, _
            "Changement de client", 48
        
    End If
    
    
    
    ' NUMERO D'ARTICLE
    ' Si la case est vide, j'insère le numéro d'article. Si l'article est déjà enregistré, j'ajoute la quantité
    ' déjà présente et l'ancienne quantité
    
    For Each article_num In Range("L22:L100")
        If IsEmpty(article_num.Value) Then
            article_num.Value = Range("C26").Value
            Exit For
        ElseIf Not IsEmpty(article_num.Value) And article_num.Value = Range("C26").Value Then
            article_num.Offset(ColumnOffset:=2).Value = article_num.Offset(ColumnOffset:=2).Value + Range("G26").Value
            Exit Sub
        End If
     
    Next article_num
    
    
    ' NOM DE L'ARTICLE
    ' Avec une boucle, j'insère le nom de l'article si la case est vide, puis je sors de la boucle
    
    For Each article_name In Range("M21:M100")
        If IsEmpty(article_name.Value) Then
            article_name.Value = Range("E26").Value
            Exit For
     
        End If
     
    Next article_name
    
    
    ' QUANTITE
    ' Même procédé que pour le nom de l'article
    
    For Each qty In Range("N21:N100")
        If IsEmpty(qty.Value) Then
            qty.Value = Range("G26").Value
            Exit For
        End If
     
    Next qty
      
End Sub
        
   
Private Sub CommandButton2_Click()
    
    ' BOUTON D'EXPORT AU FORMAT CSV
    
    ' Déclaration des variables.
    Dim content As String
    Dim rng As Range
    Dim invoice As String
    Dim customer As String
    
    
    ' Si la case L22 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("L22").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    
    ' Numéro de facture et numéro client qui vont servir pour le nom du fichier
    doc_number = Range("J22").Value
    customer = Range("K22").Value
    
    ' Je loop à travers chaque cellule qui contient les articles enregistrés afin de composer "manuellement"
    ' une variable qui contient la valeur de chaque cellule suivie d'une virgule, puis je reviens à la ligne
    ' à chaque fois que je rencontre la colonne 15, qui est la dernière valeur
    
    For Each cell In Range("L22:O24")
        If Not IsEmpty(cell.Value) Then
            content = content & cell.Value & ";"
            If cell.Column = 15 Then
                content = Left(content, Len(content) - 1) & vbNewLine
            End If
        ElseIf IsEmpty(cell.Value) Then
            Exit For
        End If
    Next cell
    
    
    ' Je crée ensuite un objet d'écriture système
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set wfile = FSO.CreateTextFile(Application.ThisWorkbook.Path & "/Panier_" & doc_number & "_Client_" & customer & ".csv", 2)
    
    ' Sur lequel j'écris le contenu de la variable que j'ai créée, puis j'enregistre
    wfile.WriteLine content
    wfile.Close
    
    ' Je confirme la création du fichier à l'utilisateur
    CreateObject("WScript.Shell").PopUp "Fichier CSV créé.", 3, _
            "Succès", 0
End Sub

Private Sub CommandButton3_Click()
    
    ' BOUTON DE REINITIALISATION DU PANIER
    Range("J22:N26").ClearContents
    
End Sub

Private Sub CommandButton4_Click()

    ' GENERATION DE FACTURE AU FORMAT PDF
    
    ' Si la case L22 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("L22").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    ' Je déclare le sheet template en temps que variabe pour travailler avec
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Template")
    
    ' Je modifie le template en fonction des informations du panier
    
    ' Type de doc
    ws.Range("F1") = "FACTURE"
    ' Numéro
    ws.Range("F2") = Range("J22")
    ' Date
    ws.Range("F3") = Range("C16")
    ' Infos client
    ws.Range("G5:I9") = Range("F15:I21").Value
    
    ' Ajouter les articles du panier à la facture
    For Each cell In Range("L22:O24")
        If Not IsEmpty(cell.Value) Then
            content = content & cell.Value & " "
            If cell.Column = 15 Then
                content = Left(content, Len(content) - 1) & vbNewLine
            End If
        ElseIf IsEmpty(cell.Value) Then
            Exit For
        End If
    Next cell
    
    MsgBox content
    
    
    'Save Active Sheet(s) as PDF
    'Sheets("Template").ExportAsFixedFormat Type:=xlTypePDF, _
    'Filename:=Application.ThisWorkbook.Path & "/Facture.pdf"
    
    ' Je confirme la création du fichier à l'utilisateur
    CreateObject("WScript.Shell").PopUp "Facture créée avec succès.", 3, _
            "Succès", 0



End Sub
