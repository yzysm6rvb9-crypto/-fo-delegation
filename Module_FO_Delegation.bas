Option Explicit

Private Const CODE_APP As String = "CMC"
Private Const DEST_FO As String = "focastmetal@gmail.com"

Sub OuvrirCalendrier()
    Dim prenom As String, code As String
    prenom = Trim(Sheets("Connexion").Range("B3").Value)
    code = Trim(Sheets("Connexion").Range("B4").Value)

    If prenom = "" Then
        MsgBox "Choisis ton prénom.", vbExclamation
        Exit Sub
    End If

    If UCase(code) <> CODE_APP Then
        MsgBox "Code incorrect.", vbCritical
        Exit Sub
    End If

    Sheets("Calendrier").Visible = xlSheetVisible
    Sheets("Mail FO").Visible = xlSheetVisible
    Sheets("Calendrier").Activate
    Range("A1").Select
End Sub

Sub EnvoyerBoiteMailFO()
    Dim wsC As Worksheet, wsA As Worksheet
    Dim prenom As String, moisTxt As String, destinataire As String
    Dim sujet As String, corps As String, detail As String
    Dim r As Long, lastA As Long
    Dim totalDeleg As Double, droit As Double, reste As Double
    Dim olApp As Object, olMail As Object

    Set wsC = Sheets("Calendrier")
    Set wsA = Sheets("Archives")

    prenom = Sheets("Connexion").Range("B3").Value
    moisTxt = Format(DateSerial(Sheets("Connexion").Range("B6").Value, Sheets("Connexion").Range("B5").Value, 1), "mmmm yyyy")
    destinataire = DEST_FO
    droit = Sheets("Connexion").Range("B7").Value
    totalDeleg = wsC.Range("E4").Value
    reste = wsC.Range("H4").Value

    sujet = "Feuille heures délégation FO - " & prenom & " - " & moisTxt
    corps = "Bonjour," & vbCrLf & vbCrLf
    corps = corps & "Veuillez trouver ci-dessous le détail des heures de délégation pour " & prenom & " - " & moisTxt & "." & vbCrLf & vbCrLf
    corps = corps & "Droit mensuel : " & Format(droit, "0.00") & " h" & vbCrLf
    corps = corps & "Total délégation prise : " & Format(totalDeleg, "0.00") & " h" & vbCrLf
    corps = corps & "Reste disponible : " & Format(reste, "0.00") & " h" & vbCrLf & vbCrLf
    corps = corps & "Détail des journées avec délégation :" & vbCrLf
    corps = corps & "------------------------------------" & vbCrLf

    detail = ""
    For r = 12 To 42
        If Val(wsC.Cells(r, "F").Value) > 0 Then
            detail = detail & Format(wsC.Cells(r, "A").Value, "dd/mm/yyyy") & " - " & wsC.Cells(r, "B").Value & _
                " | Horaire : " & wsC.Cells(r, "D").Value & _
                " | Total journée : " & Format(wsC.Cells(r, "E").Value, "0.00") & " h" & _
                " | Délégation : " & Format(wsC.Cells(r, "F").Value, "0.00") & " h" & _
                " | Travail réel : " & Format(wsC.Cells(r, "G").Value, "0.00") & " h" & _
                " | Statut : " & wsC.Cells(r, "H").Value
            If wsC.Cells(r, "I").Value <> "" Then detail = detail & " | Commentaire : " & wsC.Cells(r, "I").Value
            detail = detail & vbCrLf
        End If
    Next r

    If detail = "" Then detail = "Aucune heure de délégation déclarée sur ce mois." & vbCrLf
    corps = corps & detail & vbCrLf & "Cordialement," & vbCrLf & prenom

    lastA = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row + 1
    If lastA < 4 Then lastA = 4
    wsA.Cells(lastA, 1).Value = Now
    wsA.Cells(lastA, 2).Value = prenom
    wsA.Cells(lastA, 3).Value = Sheets("Connexion").Range("B5").Value
    wsA.Cells(lastA, 4).Value = Sheets("Connexion").Range("B6").Value
    wsA.Cells(lastA, 5).Value = droit
    wsA.Cells(lastA, 6).Value = totalDeleg
    wsA.Cells(lastA, 7).Value = reste
    wsA.Cells(lastA, 8).Value = destinataire
    wsA.Cells(lastA, 9).Value = detail

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If Not olApp Is Nothing Then
        Set olMail = olApp.CreateItem(0)
        With olMail
            .To = destinataire
            .Subject = sujet
            .Body = corps
            .Display
        End With
    Else
        MsgBox "Archive créée. Outlook n'est pas disponible sur cet appareil. Copie le texte depuis la feuille 'Mail FO'.", vbInformation
        Sheets("Mail FO").Activate
    End If
End Sub

Sub ReinitialiserMois()
    If MsgBox("Effacer les saisies du mois ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    With Sheets("Calendrier")
        .Range("D12:D42").ClearContents
        .Range("F12:F42").ClearContents
        .Range("I12:I42").ClearContents
    End With
    MsgBox "Mois réinitialisé.", vbInformation
End Sub