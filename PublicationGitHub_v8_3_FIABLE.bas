Option Explicit

Private Const GITHUB_OWNER As String = "Ploutche"
Private Const GITHUB_REPO As String = "planning-tridel"
Private Const GITHUB_BRANCH As String = "main"
Private Const GITHUB_FILE As String = "planning.json"

Private Const CONFIG_SHEET As String = "_CONFIG_GITHUB"
Private Const DATE_ROW As Long = 5
Private Const FIRST_EMPLOYEE_ROW As Long = 8
Private Const NAME_COL As Long = 3
Private Const FIRST_DATE_COL As Long = 5

Public Sub PublierPlanning()
    On Error GoTo GestionErreur

    Application.ScreenUpdating = False
    Application.StatusBar = "Vérification du planning..."
    ThisWorkbook.Save

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)

    Dim nbEmployes As Long
    Dim nbCaches As Long
    Dim nbDates As Long
    Dim avertissements As String

    ValiderPlanning ws, nbEmployes, nbCaches, nbDates, avertissements

    If Len(avertissements) > 0 Then
        Dim rep As VbMsgBoxResult
        rep = MsgBox( _
            "Le planning peut être publié, mais j'ai détecté :" & vbCrLf & vbCrLf & _
            avertissements & vbCrLf & _
            "Souhaites-tu publier quand même ?", _
            vbQuestion + vbYesNo, _
            "Vérification du planning")
        If rep <> vbYes Then GoTo Fin
    End If

    Dim token As String
    token = ObtenirTokenGitHub()
    If Len(token) = 0 Then GoTo Fin

    VerifierAccesDepot token

    Application.StatusBar = "Création des données..."

    Dim publicationId As String
    publicationId = CreerIdentifiantPublication()

    Dim contenuJSON As String
    contenuJSON = ConstruirePlanningJSON(ws, publicationId, nbEmployes, nbDates)

    VerifierJSONLocal contenuJSON, publicationId, nbEmployes, nbDates

    Application.StatusBar = "Publication sur GitHub..."
    EnvoyerSurGitHub contenuJSON, token

    Application.StatusBar = "Vérification de la publication..."
    VerifierPublicationDistante contenuJSON, token, publicationId

    Application.StatusBar = False

    MsgBox _
        "Planning publié ET vérifié avec succès." & vbCrLf & vbCrLf & _
        nbEmployes & " employés publiés" & vbCrLf & _
        nbCaches & " lignes masquées ignorées" & vbCrLf & _
                nbDates & " jours analysés" & vbCrLf & _
        "Identifiant : " & publicationId & vbCrLf & _
        "Mise à jour : " & Format$(Now, "dd.mm.yyyy à hh:nn"), _
        vbInformation, _
        "Publication vérifiée"

Fin:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

GestionErreur:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "Publication impossible"
End Sub

Public Sub InstallerBoutonPublication()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    On Error Resume Next
    ws.Shapes("BoutonPublierPlanning").Delete
    On Error GoTo 0

    Dim bouton As Shape
    Set bouton = ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 20, 190, 42)

    With bouton
        .Name = "BoutonPublierPlanning"
        .TextFrame2.TextRange.Text = "Publier le planning"
        .OnAction = "'" & ThisWorkbook.Name & "'!PublierPlanning"
        .Fill.ForeColor.RGB = RGB(38, 103, 245)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    MsgBox "Le bouton « Publier le planning » a été ajouté.", vbInformation
End Sub

Public Sub ChangerTokenGitHub()
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleConfiguration()
    ws.Range("B1").ClearContents
    Call ObtenirTokenGitHub
End Sub

Private Sub ValiderPlanning( _
    ByVal ws As Worksheet, _
    ByRef nbEmployes As Long, _
    ByRef nbCaches As Long, _
    ByRef nbDates As Long, _
    ByRef avertissements As String)

    If Not IsNumeric(ws.Range("A1").Value) Then
        Err.Raise vbObjectError + 300, , _
            "Erreur dans Excel : la cellule A1 doit contenir l'année du planning."
    End If

    Dim annee As Long
    annee = CLng(ws.Range("A1").Value)

    If annee < 2020 Or annee > 2100 Then
        Err.Raise vbObjectError + 301, , _
            "Erreur dans Excel : l'année indiquée en A1 semble incorrecte (" & annee & ")."
    End If

    Dim derniereColonne As Long
    derniereColonne = ws.Cells(DATE_ROW, ws.Columns.Count).End(xlToLeft).Column

    If derniereColonne < FIRST_DATE_COL Then
        Err.Raise vbObjectError + 302, , _
            "Erreur dans Excel : aucune date n'a été trouvée sur la ligne " & DATE_ROW & "."
    End If

    Dim datesVues As Object
    Set datesVues = CreateObject("Scripting.Dictionary")

    Dim c As Long
    For c = FIRST_DATE_COL To derniereColonne
        Dim valeurDate As Variant
        valeurDate = ws.Cells(DATE_ROW, c).Value

        If Len(Trim$(CStr(valeurDate))) > 0 Then
            If Not IsDate(valeurDate) Then
                Err.Raise vbObjectError + 303, , _
                    "Date invalide en " & ws.Cells(DATE_ROW, c).Address(False, False) & _
                    " : « " & CStr(valeurDate) & " »."
            End If

            Dim dateISO As String
            dateISO = Format$(CDate(valeurDate), "yyyy-mm-dd")

            ' Une date peut apparaître deux fois dans le classeur
            ' (ex. 01.03.2026 en BL5 et BM5).
            ' Ce n'est pas une erreur : seule la dernière colonne sera publiée.
            If Not datesVues.Exists(dateISO) Then
                datesVues.Add dateISO, True
                nbDates = nbDates + 1
            End If
        End If
    Next c

    If nbDates = 0 Then
        Err.Raise vbObjectError + 305, , "Aucune date valide n'a été trouvée."
    End If

    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).Row

    Dim nomsVus As Object
    Set nomsVus = CreateObject("Scripting.Dictionary")

    Dim codesInconnus As Object
    Set codesInconnus = CreateObject("Scripting.Dictionary")

    Dim lignesSansNom As Long
    Dim r As Long

    For r = FIRST_EMPLOYEE_ROW To derniereLigne
        If ws.Rows(r).Hidden Then
            nbCaches = nbCaches + 1
        Else
            Dim nom As String
            nom = NettoyerNomAffiche(Trim$(CStr(ws.Cells(r, NAME_COL).Value)))

            If Len(nom) > 0 And UCase$(nom) <> "VIDE" And Not EstPosteVacant(nom) Then
                Dim cleNom As String
                cleNom = UCase$(nom)

                If nomsVus.Exists(cleNom) Then
                    Err.Raise vbObjectError + 306, , _
                        "Nom en double parmi les lignes visibles : « " & nom & " »." & vbCrLf & _
                        "Corrige le doublon avant de publier."
                End If

                nomsVus.Add cleNom, True
                nbEmployes = nbEmployes + 1

                For c = FIRST_DATE_COL To derniereColonne
                    If IsDate(ws.Cells(DATE_ROW, c).Value) Then
                        Dim code As String
                        If Not IsError(ws.Cells(r, c).Value) Then
                            code = UCase$(Trim$(CStr(ws.Cells(r, c).Value)))
                            If Len(code) > 0 And Not CodeConnu(code) Then
                                If Not codesInconnus.Exists(code) Then codesInconnus.Add code, True
                            End If
                        End If
                    End If
                Next c
            ElseIf LigneContientPlanning(ws, r, derniereColonne) Then
                lignesSansNom = lignesSansNom + 1
            End If
        End If
    Next r

    If nbEmployes = 0 Then
        Err.Raise vbObjectError + 307, , _
            "Aucun employé visible à publier." & vbCrLf & _
            "Les lignes masquées sont volontairement ignorées."
    End If

    ' Les lignes visibles avec planning mais sans nom sont volontairement ignorées.
    ' Elles ne doivent pas bloquer la publication.
    If lignesSansNom > 0 Then
        avertissements = avertissements & _
            lignesSansNom & " ligne(s) sans nom ignorée(s)." & vbCrLf
    End If

    If codesInconnus.Count > 0 Then
        avertissements = avertissements & _
            "Code(s) non répertorié(s) publié(s) sans légende : " & _
            Join(codesInconnus.Keys, ", ") & "." & vbCrLf
    End If
End Sub

Private Function LigneContientPlanning( _
    ByVal ws As Worksheet, _
    ByVal ligne As Long, _
    ByVal derniereColonne As Long) As Boolean

    Dim c As Long
    For c = FIRST_DATE_COL To derniereColonne
        If Len(Trim$(CStr(ws.Cells(ligne, c).Value))) > 0 Then
            LigneContientPlanning = True
            Exit Function
        End If
    Next c
End Function

Private Function CodeConnu(ByVal code As String) As Boolean
    Const CODES As String = _
        "|R|RM|RA|RN|VA|VJ|HOR|HSR|PIQ|RF|CA|GR|EN|FE|PO|MI|PI|PC|EM|AN|AP|PR|MS|MC|"
    CodeConnu = (InStr(1, CODES, "|" & UCase$(code) & "|", vbTextCompare) > 0)
End Function


Private Function NettoyerNomAffiche(ByVal nom As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\s*\([^)]*\)\s*"
    regex.Global = True

    nom = regex.Replace(nom, " ")

    Do While InStr(nom, "  ") > 0
        nom = Replace(nom, "  ", " ")
    Loop

    NettoyerNomAffiche = Trim$(nom)
End Function

Private Function EstPosteVacant(ByVal nom As String) As Boolean
    Dim n As String
    n = LCase$(Trim$(NettoyerNomAffiche(nom)))

    EstPosteVacant = (n = "poste vacant" Or n = "poste-vacant" Or n = "vacant")
End Function


Private Function LireCommentaireCellule(ByVal cellule As Range) As String
    Dim texte As String
    texte = ""

    On Error Resume Next

    ' Notes / anciens commentaires Excel
    If Not cellule.Comment Is Nothing Then
        texte = Trim$(CStr(cellule.Comment.Text))
    End If

    ' Commentaires modernes : accès tardif pour rester compatible
    If Len(texte) = 0 Then
        Dim commentaireThreaded As Object
        Set commentaireThreaded = CallByName(cellule, "CommentThreaded", VbGet)

        If Not commentaireThreaded Is Nothing Then
            texte = Trim$(CStr(CallByName(commentaireThreaded, "Text", VbGet)))
        End If
    End If

    On Error GoTo 0

    LireCommentaireCellule = texte
End Function

Private Function ConstruirePlanningJSON(ByVal ws As Worksheet, ByVal publicationId As String, ByVal nbEmployes As Long, ByVal nbDates As Long) As String
    Dim annee As Long
    annee = CLng(ws.Range("A1").Value)

    Dim derniereColonne As Long
    derniereColonne = ws.Cells(DATE_ROW, ws.Columns.Count).End(xlToLeft).Column

    Dim derniereLigne As Long
    derniereLigne = ws.Cells(ws.Rows.Count, NAME_COL).End(xlUp).Row

    Dim resultat As String
    resultat = "{""schema_version"":2" & _
        ",""publication_id"":""" & EchappeJSON(publicationId) & """" & _
        ",""employee_count"":" & CStr(nbEmployes) & _
        ",""date_count"":" & CStr(nbDates) & _
        ",""annee"":" & CStr(annee) & _
        ",""mise_a_jour"":""" & Format$(Now, "yyyy-mm-dd\THH:nn:ss") & _
        """,""employes"":["

    Dim premierePersonne As Boolean
    premierePersonne = True

    Dim r As Long
    For r = FIRST_EMPLOYEE_ROW To derniereLigne

        ' IMPORTANT : toute ligne masquée dans Excel est absente du site.
        If Not ws.Rows(r).Hidden Then

            Dim nom As String
            nom = NettoyerNomAffiche(Trim$(CStr(ws.Cells(r, NAME_COL).Value)))

            If Len(nom) > 0 And UCase$(nom) <> "VIDE" And Not EstPosteVacant(nom) Then
                If Not premierePersonne Then resultat = resultat & ","
                premierePersonne = False

                ' Plus aucun matricule n'est envoyé.
                ' L'identifiant technique utilise simplement la ligne Excel.
                Dim identifiant As String
                identifiant = "ligne-" & CStr(r)

                resultat = resultat & _
                    "{""id"":""" & identifiant & _
                    """,""nom"":""" & EchappeJSON(nom) & _
                    """,""planning"":{"

                Dim premiereDate As Boolean
                premiereDate = True

                Dim c As Long
                For c = FIRST_DATE_COL To derniereColonne
                    If IsDate(ws.Cells(DATE_ROW, c).Value) Then

                        Dim dateISO As String
                        dateISO = Format$(CDate(ws.Cells(DATE_ROW, c).Value), "yyyy-mm-dd")

                        Dim ignorerColonneDoublon As Boolean
                        ignorerColonneDoublon = False

                        ' Si la colonne suivante porte exactement la même date,
                        ' on ignore la colonne actuelle et on utilise la dernière.
                        If c < derniereColonne Then
                            If IsDate(ws.Cells(DATE_ROW, c + 1).Value) Then
                                If Format$(CDate(ws.Cells(DATE_ROW, c + 1).Value), "yyyy-mm-dd") = dateISO Then
                                    ignorerColonneDoublon = True
                                End If
                            End If
                        End If

                        If Not ignorerColonneDoublon Then
                            If Not premiereDate Then resultat = resultat & ","
                            premiereDate = False

                            Dim code As String
                            If IsError(ws.Cells(r, c).Value) Or IsEmpty(ws.Cells(r, c).Value) Then
                                code = ""
                            Else
                                code = Trim$(CStr(ws.Cells(r, c).Value))
                            End If

                            resultat = resultat & """" & dateISO & """:"

                            If Len(code) = 0 Then
                                resultat = resultat & "null"
                            Else
                                resultat = resultat & """" & EchappeJSON(code) & """"
                            End If
                        End If
                    End If
                Next c

                resultat = resultat & "},""commentaires"":{"

                Dim premierCommentaire As Boolean
                premierCommentaire = True

                For c = FIRST_DATE_COL To derniereColonne
                    If IsDate(ws.Cells(DATE_ROW, c).Value) Then

                        Dim dateCommentaireISO As String
                        dateCommentaireISO = Format$(CDate(ws.Cells(DATE_ROW, c).Value), "yyyy-mm-dd")

                        Dim ignorerCommentaireDoublon As Boolean
                        ignorerCommentaireDoublon = False

                        If c < derniereColonne Then
                            If IsDate(ws.Cells(DATE_ROW, c + 1).Value) Then
                                If Format$(CDate(ws.Cells(DATE_ROW, c + 1).Value), "yyyy-mm-dd") = dateCommentaireISO Then
                                    ignorerCommentaireDoublon = True
                                End If
                            End If
                        End If

                        If Not ignorerCommentaireDoublon Then
                            Dim texteCommentaire As String
                            texteCommentaire = LireCommentaireCellule(ws.Cells(r, c))

                            If Len(texteCommentaire) > 0 Then
                                If Not premierCommentaire Then resultat = resultat & ","
                                premierCommentaire = False

                                resultat = resultat & _
                                    """" & dateCommentaireISO & """:""" & _
                                    EchappeJSON(texteCommentaire) & """"
                            End If
                        End If
                    End If
                Next c

                resultat = resultat & "}}"
            End If
        End If
    Next r

    ConstruirePlanningJSON = resultat & "]}"
End Function

Private Sub VerifierAccesDepot(ByVal token As String)
    Dim url As String
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & GITHUB_REPO

    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "GET", url, False
    AjouterEntetes req, token
    req.Send

    If req.Status = 401 Then
        Err.Raise vbObjectError + 310, , _
            "Le token GitHub est invalide ou expiré." & vbCrLf & _
            "Lance « ChangerTokenGitHub » puis colle un nouveau token."
    ElseIf req.Status = 403 Then
        Err.Raise vbObjectError + 311, , _
            "Le token n'a pas les droits nécessaires." & vbCrLf & _
            "Il doit avoir Contents: Read and write sur planning-tridel."
    ElseIf req.Status = 404 Then
        Err.Raise vbObjectError + 312, , _
            "Le token n'a pas accès au dépôt Ploutche/planning-tridel."
    ElseIf req.Status <> 200 Then
        Err.Raise vbObjectError + 313, , _
            "Erreur GitHub " & req.Status & ":" & vbCrLf & req.ResponseText
    End If
End Sub

Private Sub EnvoyerSurGitHub(ByVal contenuJSON As String, ByVal token As String)
    Dim url As String
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & _
          GITHUB_REPO & "/contents/" & GITHUB_FILE

    Dim sha As String
    sha = LireShaExistant(url, token)

    Dim corps As String
    corps = "{""message"":""Mise à jour du planning depuis Excel""," & _
            """content"":""" & Base64UTF8(contenuJSON) & """," & _
            """branch"":""" & GITHUB_BRANCH & """"

    If Len(sha) > 0 Then corps = corps & ",""sha"":""" & sha & """"
    corps = corps & "}"

    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "PUT", url, False
    AjouterEntetes req, token
    req.SetRequestHeader "Content-Type", "application/json"
    req.Send CorpsUTF8(corps)

    If req.Status <> 200 And req.Status <> 201 Then
        Err.Raise vbObjectError + 314, , _
            "GitHub a refusé la mise à jour (" & req.Status & ")." & vbCrLf & _
            req.ResponseText
    End If
End Sub

Private Function LireShaExistant(ByVal url As String, ByVal token As String) As String
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "GET", url & "?ref=" & GITHUB_BRANCH, False
    AjouterEntetes req, token
    req.Send

    If req.Status = 404 Then
        LireShaExistant = ""
        Exit Function
    End If

    If req.Status <> 200 Then
        Err.Raise vbObjectError + 315, , _
            "Impossible de vérifier planning.json (" & req.Status & ")." & vbCrLf & _
            req.ResponseText
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = """sha""\s*:\s*""([^""]+)"""
    regex.Global = False

    Dim matches As Object
    Set matches = regex.Execute(req.ResponseText)

    If matches.Count = 0 Then
        Err.Raise vbObjectError + 316, , "Le SHA de planning.json est introuvable."
    End If

    LireShaExistant = matches(0).SubMatches(0)
End Function


Private Function CreerIdentifiantPublication() As String
    Randomize
    CreerIdentifiantPublication = _
        Format$(Now, "yyyymmdd-hhnnss") & "-" & _
        Format$(CLng(Rnd() * 999999), "000000")
End Function

Private Sub VerifierJSONLocal( _
    ByVal contenuJSON As String, _
    ByVal publicationId As String, _
    ByVal nbEmployes As Long, _
    ByVal nbDates As Long)

    If Len(contenuJSON) < 100 Then
        Err.Raise vbObjectError + 320, , _
            "Le JSON généré est anormalement petit. Publication annulée."
    End If

    If InStr(1, contenuJSON, """schema_version"":2", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 321, , _
            "Le JSON généré ne contient pas la version de schéma attendue."
    End If

    If InStr(1, contenuJSON, """publication_id"":""" & publicationId & """", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 322, , _
            "L'identifiant de publication manque dans le JSON généré."
    End If

    If InStr(1, contenuJSON, """employee_count"":" & CStr(nbEmployes), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 323, , _
            "Le nombre d'employés du JSON ne correspond pas au contrôle Excel."
    End If

    If InStr(1, contenuJSON, """date_count"":" & CStr(nbDates), vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 324, , _
            "Le nombre de dates du JSON ne correspond pas au contrôle Excel."
    End If

    If Right$(contenuJSON, 2) <> "]}" Then
        Err.Raise vbObjectError + 325, , _
            "Le JSON généré semble tronqué. Publication annulée."
    End If
End Sub

Private Sub VerifierPublicationDistante( _
    ByVal contenuAttendu As String, _
    ByVal token As String, _
    ByVal publicationId As String)

    Dim tentative As Long

    For tentative = 1 To 3
        Dim contenuDistant As String
        contenuDistant = LireContenuGitHub(token)

        If contenuDistant = contenuAttendu Then
            If InStr(1, contenuDistant, """publication_id"":""" & publicationId & """", vbBinaryCompare) > 0 Then
                Exit Sub
            End If
        End If

        If tentative < 3 Then
            Application.Wait Now + TimeSerial(0, 0, 1)
        End If
    Next tentative

    Err.Raise vbObjectError + 326, , _
        "GitHub a accepté l'envoi, mais le fichier relu ne correspond pas exactement au JSON généré." & vbCrLf & _
        "La publication n'est PAS considérée comme validée." & vbCrLf & _
        "Ne communique pas cette mise à jour tant que le problème n'est pas corrigé."
End Sub

Private Function LireContenuGitHub(ByVal token As String) As String
    Dim url As String
    url = "https://api.github.com/repos/" & GITHUB_OWNER & "/" & _
          GITHUB_REPO & "/contents/" & GITHUB_FILE & _
          "?ref=" & GITHUB_BRANCH & "&v=" & Format$(Timer * 1000, "0")

    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "GET", url, False
    AjouterEntetes req, token
    req.Send

    If req.Status <> 200 Then
        Err.Raise vbObjectError + 327, , _
            "Impossible de relire planning.json après publication (" & req.Status & ")." & vbCrLf & _
            req.ResponseText
    End If

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = """content""\s*:\s*""([^""]+)"""
    regex.Global = False

    Dim matches As Object
    Set matches = regex.Execute(req.ResponseText)

    If matches.Count = 0 Then
        Err.Raise vbObjectError + 328, , _
            "GitHub n'a pas renvoyé le contenu de planning.json."
    End If

    Dim base64Texte As String
    base64Texte = matches(0).SubMatches(0)
    base64Texte = Replace(base64Texte, "\n", "")
    base64Texte = Replace(base64Texte, "\r", "")
    base64Texte = Replace(base64Texte, vbCr, "")
    base64Texte = Replace(base64Texte, vbLf, "")

    LireContenuGitHub = Base64VersUTF8(base64Texte)
End Function

Private Function Base64VersUTF8(ByVal texteBase64 As String) As String
    Dim doc As Object
    Dim noeud As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set noeud = doc.createElement("base64")

    noeud.DataType = "bin.base64"
    noeud.Text = texteBase64

    Dim octets As Variant
    octets = noeud.nodeTypedValue

    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 1
    flux.Open
    flux.Write octets
    flux.Position = 0
    flux.Type = 2
    flux.Charset = "utf-8"

    Base64VersUTF8 = flux.ReadText
    flux.Close
End Function

Private Sub AjouterEntetes(ByVal req As Object, ByVal token As String)
    req.SetRequestHeader "Accept", "application/vnd.github+json"
    req.SetRequestHeader "Authorization", "Bearer " & Trim$(token)
    req.SetRequestHeader "X-GitHub-Api-Version", "2022-11-28"
    req.SetRequestHeader "User-Agent", "Planning-Tridel-Excel"
End Sub

Private Function ObtenirTokenGitHub() As String
    Dim ws As Worksheet
    Set ws = ObtenirFeuilleConfiguration()

    Dim token As String
    token = Trim$(CStr(ws.Range("B1").Value))

    If Len(token) = 0 Then
        token = Trim$(InputBox( _
            "Colle le token GitHub autorisé à modifier :" & vbCrLf & _
            "Ploutche/planning-tridel", _
            "Configuration GitHub"))

        If Len(token) > 0 Then
            ws.Range("B1").Value = token
            ws.Visible = xlSheetVeryHidden
            ThisWorkbook.Save
        End If
    End If

    ObtenirTokenGitHub = token
End Function

Private Function ObtenirFeuilleConfiguration() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET
        ws.Range("A1").Value = "Token GitHub"
        ws.Visible = xlSheetVeryHidden
    End If

    Set ObtenirFeuilleConfiguration = ws
End Function

Private Function EchappeJSON(ByVal valeur As String) As String
    valeur = Replace(valeur, "\", "\\")
    valeur = Replace(valeur, """", "\""")
    valeur = Replace(valeur, vbCrLf, "\n")
    valeur = Replace(valeur, vbCr, "\n")
    valeur = Replace(valeur, vbLf, "\n")
    valeur = Replace(valeur, vbTab, "\t")
    EchappeJSON = valeur
End Function

Private Function Base64UTF8(ByVal texte As String) As String
    Dim octets As Variant
    octets = CorpsUTF8(texte)

    Dim doc As Object
    Dim noeud As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    Set noeud = doc.createElement("base64")

    noeud.DataType = "bin.base64"
    noeud.nodeTypedValue = octets

    Base64UTF8 = Replace(Replace(noeud.Text, vbCr, ""), vbLf, "")
End Function

Private Function CorpsUTF8(ByVal texte As String) As Variant
    Dim flux As Object
    Set flux = CreateObject("ADODB.Stream")

    flux.Type = 2
    flux.Charset = "utf-8"
    flux.Open
    flux.WriteText texte
    flux.Position = 0
    flux.Type = 1
    flux.Position = 3

    CorpsUTF8 = flux.Read
    flux.Close
End Function
