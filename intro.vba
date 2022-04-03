Private Sub CommandButton1_Click()

    ' EMPECHER D'ECRIRE DANS LES CELLULES NON PREVUES
    
    ' AJOUTER ARTICLE AU PANIER
    
    ' Je commence par insérer la date dans la cellule dédiée
    Range("C11") = Date
    
    
    ' Pour chaque cellule de la liste, je vérifie si elle n'est pas vide, si sa valeur est bien numérique et si elle est positive
    ' Si elle est vide, non numérique ou négative, je change sa couleur, je notifie l'utilisateur et j'arrête le script
    
    For Each mycell In Array(Range("C13"), Range("C14"), Range("C18"), Range("C20"))
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
            mycell.Interior.ColorIndex = 6
        End If
             
    Next mycell
    
    
    ' NUMERO DOCUMENT
    ' J'ajoute le numéro de facture dans la case dédiée si la case est vide
    ' J'empêche l'utilisateur de changer de numéro de facture avec un panier rempli
    
    If IsEmpty(Range("L12").Value) Then
        Range("L12").Value = CInt(Range("C20").Value)
    ElseIf Not IsEmpty(Range("L12").Value) And Range("L12").Value <> Range("C20").Value Then
        CreateObject("WScript.Shell").PopUp "Avant d'éditer un autre document, vous devez exporter le panier client puis " & _
        "le réinitialiser.", 10, _
            "Document non finalisé", 48
        
    End If
    
    
    ' NUMERO CLIENT
    ' Si la case est vide, j'insère le numéro client
    ' Si elle n'est pas vide, j'empêche l'utilisateur de changer de client si des articles sont dans le panier
    
    If IsEmpty(Range("M12").Value) Then
        Range("M12").Value = Range("C13").Value
    ElseIf Not IsEmpty(Range("M12").Value) And Range("M12").Value <> Range("C13").Value Then
        CreateObject("WScript.Shell").PopUp "Avant de modifier le numéro du client, vous devez exporter son panier puis " & _
        "le réinitialiser.", 10, _
            "Changement de client", 48
        
    End If
    
    
    
    ' NUMERO D'ARTICLE
    ' Si la case est vide, j'insère le numéro d'article. Si l'article est déjà enregistré, j'ajoute la quantité
    ' déjà présente et l'ancienne quantité
    
    For Each article_num In Range("N12:N17")
        If IsEmpty(article_num.Value) Then
            article_num.Value = Range("C14").Value
            Exit For
        ElseIf Not IsEmpty(article_num.Value) And article_num.Value = Range("C14").Value Then
            article_num.Offset(ColumnOffset:=2).Value = article_num.Offset(ColumnOffset:=2).Value + Range("C18").Value
            Exit Sub
        End If
     
    Next article_num
    
    
    ' NOM DE L'ARTICLE
    ' Avec une boucle, j'insère le nom de l'article si la case est vide, puis je sors de la boucle
    
    For Each article_name In Range("O12:O17")
        If IsEmpty(article_name.Value) Then
            article_name.Value = Range("C15").Value
            Exit For
     
        End If
     
    Next article_name
    
    
    ' QUANTITE
    ' Même procédé que pour le nom de l'article
    
    For Each qty In Range("P12:P17")
        If IsEmpty(qty.Value) Then
            qty.Value = Range("C18").Value
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
    
    
    ' Si la case N12 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("N12").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    
    ' Numéro de facture et numéro client qui vont servir pour le nom du fichier
    doc_number = Range("L12").Value
    customer = Range("M12").Value
    
    ' Je loop à travers chaque cellule qui contient les articles enregistrés afin de composer "manuellement"
    ' une variable qui contient la valeur de chaque cellule suivie d'une virgule, puis je reviens à la ligne
    ' à chaque fois que je rencontre la colonne 15, qui est la dernière valeur
    
    For Each cell In Range("O12:Q17")
        If Not IsEmpty(cell.Value) Then
            content = content & cell.Value & ";"
            If cell.Column = 17 Then
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
    Range("L12:P17").ClearContents
    
End Sub

Private Sub CommandButton4_Click()

    ' GENERATION DE FACTURE AU FORMAT PDF
    
    
    ' Si la case N12 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("N12").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    ' Je déclare le sheet template en temps que variabe pour travailler avec
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Template")
    
    ' Je modifie le template en fonction des informations du panier
    
    ' Type de doc
    ws.Range("H9:H11") = "FACTURE"
    ' Date
    ws.Range("H13") = Range("C11")
    ' Numéro
    ws.Range("H15") = Range("L12")
    ' Infos client
    ws.Range("C19:E23") = Range("F11:I17").Value
    ws.Range("E17") = Range("C13")
    
    ' Ajouter les articles du panier à la facture
    
    ws.Range("C26:E27") = Range("O12")
    ws.Range("F26:F27") = Range("P12")
    ws.Range("H26:H27") = Range("Q12")
    
    ws.Range("C28:E29") = Range("O13")
    ws.Range("F28:F29") = Range("P13")
    ws.Range("H28:H29") = Range("Q13")
        
    ws.Range("C30:E31") = Range("O14")
    ws.Range("F30:F31") = Range("P14")
    ws.Range("H30:H31") = Range("Q14")
    
    ws.Range("C32:E33") = Range("O15")
    ws.Range("F32:F33") = Range("P15")
    ws.Range("H32:H33") = Range("Q15")
    
    ws.Range("C34:E35") = Range("O16")
    ws.Range("F34:F35") = Range("P16")
    ws.Range("H34:H35") = Range("Q16")
    
    ws.Range("C36:E37") = Range("O17")
    ws.Range("F36:F37") = Range("P17")
    ws.Range("H36:H37") = Range("Q17")
        
    ' Exporter le template complété au format PDF
    Sheets("Template").ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=Application.ThisWorkbook.Path & "/Facture.pdf"
    
    ' Réinitialiser le template
    ws.Range("C26:F37") = ""
    ws.Range("H26:H37") = ""
    ws.Range("C19:E23") = ""
    ws.Range("H9:H11") = ""
    ws.Range("H13") = ""
    ws.Range("H15") = ""
    ws.Range("E17") = ""
    
    ' Je confirme la création du fichier à l'utilisateur
    CreateObject("WScript.Shell").PopUp "Facture créée avec succès.", 3, _
            "Succès", 0
            
End Sub

Private Sub CommandButton5_Click()

' GENERATION DE DEVIS AU FORMAT PDF
    
    
    ' Si la case N12 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("N12").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    ' Je déclare le sheet template en temps que variabe pour travailler avec
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Template")
    
    ' Je modifie le template en fonction des informations du panier
    
    ' Type de doc
    ws.Range("H9:H11") = "DEVIS"
    ' Date
    ws.Range("H13") = Range("C11")
    ' Numéro
    ws.Range("H15") = Range("L12")
    ' Infos client
    ws.Range("C19:E23") = Range("F11:I17").Value
    ws.Range("E17") = Range("C13")
    
    ' Ajouter les articles du panier à la facture
    
    ws.Range("C26:E27") = Range("O12")
    ws.Range("F26:F27") = Range("P12")
    ws.Range("H26:H27") = Range("Q12")
    
    ws.Range("C28:E29") = Range("O13")
    ws.Range("F28:F29") = Range("P13")
    ws.Range("H28:H29") = Range("Q13")
        
    ws.Range("C30:E31") = Range("O14")
    ws.Range("F30:F31") = Range("P14")
    ws.Range("H30:H31") = Range("Q14")
    
    ws.Range("C32:E33") = Range("O15")
    ws.Range("F32:F33") = Range("P15")
    ws.Range("H32:H33") = Range("Q15")
    
    ws.Range("C34:E35") = Range("O16")
    ws.Range("F34:F35") = Range("P16")
    ws.Range("H34:H35") = Range("Q16")
    
    ws.Range("C36:E37") = Range("O17")
    ws.Range("F36:F37") = Range("P17")
    ws.Range("H36:H37") = Range("Q17")
        
    ' Exporter le template complété au format PDF
    Sheets("Template").ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=Application.ThisWorkbook.Path & "/Devis.pdf"
    
    ' Réinitialiser le template
    ws.Range("C26:F37") = ""
    ws.Range("H26:H37") = ""
    ws.Range("C19:E23") = ""
    ws.Range("H9:H11") = ""
    ws.Range("H13") = ""
    ws.Range("H15") = ""
    ws.Range("E17") = ""
    
    ' Je confirme la création du fichier à l'utilisateur
    CreateObject("WScript.Shell").PopUp "Devis créé avec succès.", 3, _
            "Succès", 0

End Sub

Private Sub CommandButton6_Click()


' GENERATION DE BON DE COMMANDE AU FORMAT PDF
    
    
    ' Si la case N12 (premier article) est vide, je demande d'enregistrer un article et j'arrête l'opération
    
    If IsEmpty(Range("N12").Value) Then
        CreateObject("WScript.Shell").PopUp "Aucun article ne figure dans le panier.", 3, "Panier vide", 0
        Exit Sub
    End If
    
    ' Je déclare le sheet template en temps que variabe pour travailler avec
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Template")
    
    ' Je modifie le template en fonction des informations du panier
    
    ' Type de doc
    ws.Range("H9:H11") = "BON DE COMMANDE"
    ' Date
    ws.Range("H13") = Range("C11")
    ' Numéro
    ws.Range("H15") = Range("L12")
    ' Infos client
    ws.Range("C19:E23") = Range("F11:I17").Value
    ws.Range("E17") = Range("C13")
    
    ' Ajouter les articles du panier à la facture
    
    ws.Range("C26:E27") = Range("O12")
    ws.Range("F26:F27") = Range("P12")
    ws.Range("H26:H27") = Range("Q12")
    
    ws.Range("C28:E29") = Range("O13")
    ws.Range("F28:F29") = Range("P13")
    ws.Range("H28:H29") = Range("Q13")
        
    ws.Range("C30:E31") = Range("O14")
    ws.Range("F30:F31") = Range("P14")
    ws.Range("H30:H31") = Range("Q14")
    
    ws.Range("C32:E33") = Range("O15")
    ws.Range("F32:F33") = Range("P15")
    ws.Range("H32:H33") = Range("Q15")
    
    ws.Range("C34:E35") = Range("O16")
    ws.Range("F34:F35") = Range("P16")
    ws.Range("H34:H35") = Range("Q16")
    
    ws.Range("C36:E37") = Range("O17")
    ws.Range("F36:F37") = Range("P17")
    ws.Range("H36:H37") = Range("Q17")
        
    ' Exporter le template complété au format PDF
    Sheets("Template").ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=Application.ThisWorkbook.Path & "/BC.pdf"
    
    ' Réinitialiser le template
    ws.Range("C26:F37") = ""
    ws.Range("H26:H37") = ""
    ws.Range("C19:E23") = ""
    ws.Range("H9:H11") = ""
    ws.Range("H13") = ""
    ws.Range("H15") = ""
    ws.Range("E17") = ""
    
    ' Je confirme la création du fichier à l'utilisateur
    CreateObject("WScript.Shell").PopUp "Bon de commande créé avec succès.", 3, _
            "Succès", 0

End Sub
