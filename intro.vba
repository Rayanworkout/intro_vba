Private Sub CommandButton1_Click()

    Dim empty_cells As Long
         
         
    ' Vérifier pour chaque cellule de la liste (Array) si celle-ci est vide
    ' Si elle est vide, changer la couleur et ajouter 1 à la variable i
    ' La variable i contient le nombre de cases vides, on l'incrémente à chaque tour de boucle
    
    For Each myCell In Array(Range("C15"), Range("C17"), Range("C25"), Range("G25"))
        If IsEmpty(myCell.Value) Then
            myCell.Interior.ColorIndex = 6
            empty_cells = empty_cells + 1
        ElseIf Not IsEmpty(myCell.Value) Then
            myCell.Interior.ColorIndex = 2
        End If
    Next myCell
    
    ' Si il n'y a qu'une case vide, envoyer le message au singulier puis arrêter le programme
    If empty_cells = 1 Then
        MsgBox _
    "Une des cases obligatoires est vide."
    Exit Sub
    
    ' Si + d'une case est vide, envoyer le message au pluriel puis arrêter le programme
    ElseIf empty_cells > 1 Then
        MsgBox _
        "" & empty_cells & " cases obligatoires sont vides."
        Exit Sub
    
    End If
    
    
    
    
    ' Après avoir vérifié que toutes les cases sont pleines, je commence à remplir la facture
    
    ' Ajouter des données aux lignes de facture grâce aux cases C15, C17, C25 et G25
    
    ' J'utilise pour ça une boucle qui passe à travers chaque ligne de la colonne L.
    ' Si la ligne est vide, j'insère valeur correspondante, sinon j'arrête la boucle
    
    
    ' DATE
    For Each mydate In Range("K21:K38")
        If IsEmpty(mydate.Value) Then
            mydate.Value = Date
            Exit For
        End If
    Next mydate
    
    ' NUMERO DE FACTURE
    For Each invoice_num In Range("L21:L38")
        If IsEmpty(invoice_num.Value) Then
            invoice_num.Value = Range("C15").Value
            Exit For
     
        End If
     
    Next invoice_num
    
    
    ' NUMERO D'ARTICLE
    For Each article_num In Range("M21:M38")
        If IsEmpty(article_num.Value) Then
            article_num.Value = Range("I27").Value
            Exit For
     
        End If
     
    Next article_num
    
    ' NOM DE L'ARTICLE
    For Each article_name In Range("N21:N38")
        If IsEmpty(article_name.Value) Then
            article_name.Value = Range("E25").Value
            Exit For
     
        End If
     
    Next article_name
    
    
    ' PRIX DE L'ARTICLE
    For Each article_price In Range("O21:O38")
        If IsEmpty(article_price.Value) Then
            article_price.Value = Range("G25").Value * Range("C26").Value  '(nombre * prix unitaire)
            Exit For
     
        End If
     
    Next article_price
    
    ' QUANTITE
    For Each qty In Range("P21:P38")
        If IsEmpty(qty.Value) Then
            qty.Value = Range("G25").Value
            Exit For
     
        End If
     
    Next qty
    
    
    ' NUMERO CLIENT
    For Each customer_num In Range("Q21:Q38")
        If IsEmpty(customer_num.Value) Then
            customer_num.Value = Range("C17").Value
            Exit For
     
        End If
     
    Next customer_num
    
    
    ' REMISE
    For Each discount In Range("R21:R38")
        If IsEmpty(discount.Value) Then
            Price = (Range("C26").Value * Range("G25").Value)
            discount.Value = Price * Range("C18").Value
            Exit For
     
        End If
     
    Next discount
    
End Sub
