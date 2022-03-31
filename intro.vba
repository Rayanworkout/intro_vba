Private Sub CommandButton1_Click()

    
    ' TODO EMPECHER DE METTRE DES INFORMATIONS ERRONEES DANS LES CASES
    
    
    Dim empty_cells As Long
    
         
    ' Vérifier pour chaque cellule de la liste (Array) si celle-ci est vide
    ' Si elle est vide, changer la couleur et ajouter 1 à la variable empty_cells
    ' La variable empty_cells contient le nombre de cases vides, on l'incrémente à chaque tour de boucle
    
    For Each myCell In Array(Range("C15"), Range("C17"), Range("C25"), Range("G25"))
        If IsEmpty(myCell.Value) Then
            myCell.Interior.ColorIndex = 3
            empty_cells = empty_cells + 1
        ElseIf Not IsEmpty(myCell.Value) Then
            myCell.Interior.ColorIndex = 2
        End If
    Next myCell
    
    ' Si il n'y a qu'une case vide, envoyer le message au singulier puis arrêter le programme
    If empty_cells = 1 Then
        CreateObject("WScript.Shell").PopUp "UNE CASE OBLIGATOIRE EST VIDE.", 2, "Cases manquantes", 0
    Exit Sub
    
    ' Si + d'une case est vide, envoyer le message au pluriel puis arrêter le programme.
    ElseIf empty_cells > 1 Then
        CreateObject("WScript.Shell").PopUp "" & empty_cells & " CASES OBLIGATOIRES SONT VIDES.", 2, "Cases manquantes", 0
        Exit Sub
    
    End If
    
    
    ' Après avoir vérifié que toutes les cases sont pleines, je remplis la facture.
    
    ' J'ajoute des données sur la ligne correspondant à chaque facture grâce aux informations de la feuille
    
    ' J'utilise pour ça une boucle qui passe à travers chaque ligne de la colonne L.
    ' Si la ligne est vide, j'insère la valeur correspondante, sinon j'arrête la boucle.
    
    
    ' NUMERO DE FACTURE
    For Each invoice_num In Range("L21:L100")
        If invoice_num.Value = Range("C15") Then  'JE VERIFIE QUE LE NUMERO DE FACTURE N'EXISTE PAS DEJA
            CreateObject("WScript.Shell").PopUp "Ce numéro de facture existe déjà !" & Chr(10) & "Merci d'en choisir un autre.", 3, "Facture existante", 0
            Exit Sub
        End If
        
        If IsEmpty(invoice_num.Value) Then
            invoice_num.Value = Range("C15").Value
            Exit For
     
        End If
     
    Next invoice_num
    
    ' DATE
    For Each mydate In Range("K21:K100")
        If IsEmpty(mydate.Value) Then
            mydate.Value = Date
            Exit For
        End If
    Next mydate
    
    
    ' NUMERO D'ARTICLE
    For Each article_num In Range("M21:M100")
        If IsEmpty(article_num.Value) Then
            article_num.Value = Range("I27").Value
            Exit For
     
        End If
     
    Next article_num
    
    ' NOM DE L'ARTICLE
    For Each article_name In Range("N21:N100")
        If IsEmpty(article_name.Value) Then
            article_name.Value = Range("E25").Value
            Exit For
     
        End If
     
    Next article_name
    
    
    ' PRIX DE L'ARTICLE
    For Each article_price In Range("O21:O100")
        If IsEmpty(article_price.Value) Then
            article_price.Value = Range("G25").Value * Range("C26").Value  '(nombre * prix unitaire)
            Exit For
     
        End If
     
    Next article_price
    
    ' QUANTITE
    For Each qty In Range("P21:P100")
        If IsEmpty(qty.Value) Then
            qty.Value = Range("G25").Value
            Exit For
     
        End If
     
    Next qty
    
    
    ' NUMERO CLIENT
    For Each customer_num In Range("Q21:Q100")
        If IsEmpty(customer_num.Value) Then
            customer_num.Value = Range("C17").Value
            Exit For
     
        End If
     
    Next customer_num
    
    
    ' REMISE
    For Each discount In Range("R21:R100")
        If IsEmpty(discount.Value) Then
            Price = (Range("C26").Value * Range("G25").Value)
            discount.Value = Price * Range("C18").Value
            Exit For
     
        End If
     
    Next discount
    
End Sub
        
   
Private Sub CommandButton2_Click()
    
    ' BOUTON POUR EXPORTER LES FACTURES AU FORMAT CSV
    
    Dim sheetExists As Boolean
    
    ' Si la case K21 (première date) est vide, je demande d'enregistrer une première facture et j'arrête l'opération
    
    If IsEmpty(Range("K21").Value) Then
        CreateObject("WScript.Shell").PopUp "Enregistrez d'abord une facture dans le panier.", 1, "Panier vide", 0
        Exit Sub
    End If
    
    ' Je vérifie si le
    For Each Sheet In Worksheets
        If Sheet.Name = "Facture" Then
            sheetExists = True
            Exit For
        End If
    Next Sheet
    
    If Not sheetExists Then
        Sheets.Add(After:=Sheets("Stocks")).Name = "Facture"
    End If
    
         
    Worksheets("Facture").Range("A1:H15") = Range("K21:R35").Value
    
    Application.DisplayAlerts = False
    
    ThisWorkbook.Sheets("Facture").Copy
    ActiveWorkbook.SaveAs Filename:=Application.ThisWorkbook.Path & "/Facture.csv", _
                          FileFormat:=xlCSV, _
                          CreateBackup:=False
    ActiveWorkbook.Close
    
    
    Worksheets("Facture").Delete
    
    Application.DisplayAlerts = True
    
    CreateObject("WScript.Shell").PopUp "Fichier .CSV créé.", 1, "Succès", 0
     
End Sub
