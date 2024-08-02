
Option Explicit On
Imports Microsoft.Office.Interop.Excel
'Imports Windows.Win32.System
Imports Excel = Microsoft.Office.Interop.Excel

Module Program

    'D�claration des variables  Excel
    Public appExcel As Excel.Application     'Application Excel
    Public wbExcel As Excel.Workbook         'Classeur Excel
    Public wsExcel As Excel.Worksheet        'Feuille Excel
    Public StrInput As String                'nom complet du fichier de donn�es
    ' dimensions des variables (� changer en cas de besoin)
    Public Const MaxFamilles As Integer = 150
    Public Const MaxPrix As Integer = 300
    Public Const MaxDenrees As Integer = 100
    ' constantes de poids
    Public Const CoefPrepa As Integer = 4
    Public Const CoefSalade As Integer = 4
    ' variables de travail
    Public Decal As Integer                 'd�calage de colonnes dans RESULTATS, 
    Public nbReport As Integer              'compteur de lignes dans RAPPORT
    Public TexteMsg As String               'variable de texte

    'D�claration variables FAMILLES

    Public NumFamille(MaxFamilles) As Integer       'Numero de caisse
    Public NomFamille(MaxFamilles) As String        'Nom de famille
    Public NBenef(MaxFamilles) As Integer           'Nombre de b�n�ficiaires
    Public SansCochon(MaxFamilles) As Boolean       'indicateur Sans cochon
    Public SansViande(MaxFamilles) As Boolean       'indicateur vegan
    Public Panier(MaxFamilles) As Single            'panier de la famille
    Public PoidsTheo(MaxFamilles) As Single         'poids th�orique de viande
    Public PanierZeu(MaxFamilles) As Single         'panier d'oeufs
    Public PoidsTheozeu(MaxFamilles) As Single      'quantit� theo d'oeufs
    Public TestSCSV(MaxFamilles) As Integer         'cocatenation des indicateurs sanscochon et sansviande
    Public NbFamilles As Integer                    'nombre de familles

    Public NbCat As Integer                         'nombre de cat�gories AIDA
    Public CheminBureau As String                   'chemin pour sauvegarde de l'image code-barre

    Sub Main()
        Dim StrOption As String
        Dim TestOption As Boolean
        Console.WriteLine("  D I S T R I B U T I O N    C R O I X - R O U G E")
        Console.WriteLine("******************************************************")
        Console.WriteLine("Donner le chemin reseau du fichier")
        StrInput = Console.ReadLine()

        'Ouverture de l'application Excel
        appExcel = CreateObject("Excel.Application")
        wbExcel = appExcel.Workbooks.Open(StrInput)
        appExcel.Visible = False

        If FeuilleExiste("RAPPORT") = True Then
            appExcel.Application.DisplayAlerts = False
            wbExcel.Worksheets("RAPPORT").delete()        'supprime la feuille avant de commencer
            appExcel.DisplayAlerts = True
        End If
        wsExcel = wbExcel.Worksheets.Add()
        wsExcel.Name = "RAPPORT"

        wsExcel.Cells(1, 1).value = "FEUILLE"
        wsExcel.Cells(1, 2).value = "Probl�me"
        wsExcel.Cells(1, 8).value = "Criticit�"
        With wsExcel.Range("A1:H1")
            .Interior.ColorIndex = 6
            .Font.Bold = True
            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End With

        Console.WriteLine("   ")
        Console.WriteLine("Ouverture du fichier Excel")
        Console.WriteLine("Choisir l'option de calcul:")
        Console.WriteLine("R�partition: tapez 1")
        Console.WriteLine("Mise � jour: tapez 2")
        Console.WriteLine("AIDA       : tapez 3")
        StrOption = Console.ReadLine()
        TestOption = False

        nbReport = 1

        Do Until TestOption = True
            Select Case StrOption
                Case "1"
                    TestOption = True
                    Call Repartition()
                Case "2"
                    TestOption = True
                    Call MAJ()
                Case "3"
                    TestOption = True
                    Call AIDA()
                Case "Exit"
                    TestOption = True
                    Call Colexit()
                Case Else
                    Console.WriteLine("Option non reconnue, tapez 1, 2 ou 3")
                    Console.WriteLine("Pour arr�ter, tapez Exit")
                    StrOption = Console.ReadLine()
            End Select
        Loop

        Console.WriteLine("Fermeture Excel, au revoir")
    End Sub

    Public Sub Colexit()
        ' ------- Sauvegarde et fermeture d'Excel -------------------------------------
        wbExcel.Save()
        wbExcel.Close() 'Fermeture du classeur Excel
        appExcel.Quit() 'Fermeture de l'application Excel
        appExcel = Nothing
        wbExcel = Nothing

        For Each ObjPro As Process In Process.GetProcessesByName("EXCEL")
            ObjPro.Kill()
        Next

    End Sub

    Private Sub Repartition()
        '**************************************
        ' CALCUL REPARTITION
        '***************************************
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim TestPrepa As String
        Dim TestPanier As Single
        Dim AlphaColTri As String
        Dim AlphaColTri2 As String
        Dim Col1 As String
        Dim Col2 As String
        Dim Col3 As String
        Dim Mode1 As Integer
        Dim Mode2 As Integer
        Dim Mode3 As Integer
        Dim A1 As New Random
        Dim ParamEcart As Integer
        Dim Saut As Boolean
        Dim Ecart As Single
        Dim EcartMaxi As Single
        Dim NumeMaxi As Integer
        Dim NbErreur As Integer

        '------------Viandes ----------------------------
        Dim ModuleViande As Single
        Dim Description(MaxDenrees) As String
        Dim Poids(MaxDenrees) As Single
        Dim Quant(MaxDenrees) As Integer
        Dim ViandeSC(MaxDenrees) As Boolean
        Dim ViandeSV(MaxDenrees) As Boolean
        Dim ResteQuant(MaxDenrees) As Integer
        Dim CPViande As String
        Dim NbDenrees As Integer
        Dim PTotViande As Single
        Dim NbTotViande As Integer
        Dim PoidsTest As String
        '----------Preparations-------------------------------
        Dim NbPreparations As Integer
        Dim Preparation(MaxDenrees) As String
        Dim TaillePrepa(MaxDenrees) As String
        Dim PoidsPrepa(MaxDenrees) As Single
        Dim QuantPrepa(MaxDenrees) As Integer
        Dim PrepaSC(MaxDenrees) As Boolean
        Dim PrepaSV(MaxDenrees) As Boolean
        Dim PTotPrepa As Single
        Dim QuantTotPrepa As Integer

        '----------Salades-------------------------------------
        Dim NbSalades As Integer
        Dim Salade(MaxDenrees) As String
        Dim TailleSalade(MaxDenrees) As String
        Dim PoidsSalade(MaxDenrees) As Single
        Dim QuantSalade(MaxDenrees) As Integer
        Dim SaladeSC(MaxDenrees) As Boolean
        Dim SaladeSV(MaxDenrees) As Boolean
        Dim PTotSalad As Single
        Dim QuantTotSalad As Integer

        '---------------Laitages-----------------------------
        Dim NbLaitages As Integer
        Dim Laitage(MaxDenrees) As String
        Dim QuantLait(MaxDenrees) As Integer
        Dim CatLait(MaxDenrees) As String
        Dim Equiv(MaxDenrees) As Single
        Dim CPLait As String
        Dim PtotLait As Single
        Dim PtotZeu As Single
        Dim CibleLait As Single
        Dim SommeLait As Single
        Dim SommeZeu As Single

        '---------D�claration variables PRIX------------------------
        Dim LibellePrix(MaxPrix) As String
        Dim CodePrix(MaxPrix) As String
        Dim CodeAida(MaxPrix) As String
        Dim NbPrix As Integer

        '---------Onglet DIVERS--------------------------------------
        Dim NbDivers As Integer
        Dim Divers(MaxDenrees) As String

        '**************************************************************************************
        '   Lecture des prix
        '**************************************************************************************
        NbErreur = 0
        If FeuilleExiste("PRIX") = False Then           'TEST pr�sence feuille
            Call Reporting("PRIX", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        wsExcel = wbExcel.Worksheets("PRIX")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"     'Compte le nbre de lignes
        NbPrix = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbPrix > MaxPrix Then            'Test d�passement dimension maxi
            Call Reporting("PRIX", "ARRET", "Nombre de prix d�passe la dimension > " & MaxPrix, "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        If NbPrix > 0 Then                  'D�but de lecture des donn�es 
            For i = 1 To NbPrix
                LibellePrix(i) = wsExcel.Cells(i + 1, 1).value
                CodePrix(i) = wsExcel.Cells(i + 1, 2).value
                CodeAida(i) = wsExcel.Cells(i + 1, 3).value
                If i > 1 Then
                    For j = 1 To i - 1          'Test codes en doublon
                        If CodePrix(i) = CodePrix(j) Then
                            TexteMsg = "Colonne B, lignes " & i + 1 & " et " & j + 1 & " CodePrix " & CodePrix(i) & " en doublon"
                            Call Reporting("PRIX", "ALERTE", TexteMsg, "PRIX")
                            NbErreur += 1
                        End If
                    Next j
                End If
            Next i
        Else
            TexteMsg = "Pas de code prix document�"
            Call Reporting("PRIX", "ALERTE", TexteMsg, "PRIX")
        End If

        '*****************************************************************
        '   Lecture des viandes
        '******************************************************************

        If FeuilleExiste("VIANDES") = False Then
            Call Reporting("VIANDES", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("VIANDES")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbDenrees = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbDenrees > MaxDenrees Then
            Call Reporting("VIANDES", "ARRET", "Nombre de viandes d�passe la dimension > " & MaxDenrees, "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        Col1 = "B"      'Pr�paration du tri des donn�es Viandes
        Mode1 = 2       'd�croissant
        Col2 = ""
        Mode2 = 1
        Col3 = ""
        Mode3 = 1
        Call TriMultiple("VIANDES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 6, NbDenrees + 1)

        PTotViande = 0
        If NbDenrees > 0 Then
            For i = 1 To NbDenrees
                Description(i) = wsExcel.Cells(i + 1, 1).value
                PoidsTest = wsExcel.Cells(i + 1, 2).value
                If VarType(wsExcel.Cells(i + 1, 2).value) = 5 Then      'teste le contenu de la cellule: vartype=5 => nombre
                    Poids(i) = wsExcel.Cells(i + 1, 2).value
                Else
                    TexteMsg = "Poids " & PoidsTest & "  � la ligne " & i + 1 & " n'est pas un nombre!"
                    Call Reporting("VIANDES", "ALERTE", TexteMsg, "VIANDES")
                    NbErreur += 1
                End If
                PoidsTest = wsExcel.Cells(i + 1, 3).value
                If VarType(wsExcel.Cells(i + 1, 3).value) = 5 Then       'teste le contenu de la cellule: vartype=5 => nombre
                    Quant(i) = wsExcel.Cells(i + 1, 3).value
                Else
                    TexteMsg = "Quantit� " & PoidsTest & "  � la ligne " & i + 1 & " n'est pas un nombre!"
                    Call Reporting("VIANDES", "ALERTE", TexteMsg, "VIANDES")
                    NbErreur += 1
                End If
                ViandeSC(i) = False                 'SC = Sans Cochon = musulman
                If wsExcel.Cells(i + 1, 4).value = 1 Then ViandeSC(i) = True
                ViandeSV(i) = False                 'SV = Sans Viande = vegan
                If wsExcel.Cells(i + 1, 5).value = 1 Then ViandeSV(i) = True
                ResteQuant(i) = Quant(i)            'initialise le Reste � la quantit� initiale
                CPViande = wsExcel.Cells(i + 1, 6).value
                PTotViande += Poids(i) * Quant(i)       'cumul du poids*quantit� avec le total
            Next
        End If

        '*****************************************************************
        '   Lecture des pr�parations
        '******************************************************************

        If FeuilleExiste("PREPARATIONS") = False Then
            Call Reporting("PREPARATIONS", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("PREPARATIONS")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbPreparations = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbPreparations > MaxDenrees Then
            Call Reporting("PREPARATIONS", "ARRET", "Nombre de preparations d�passe la dimension > " & MaxDenrees, "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        Col1 = "B"
        Mode1 = 1       'croissant
        Col2 = ""
        Mode2 = 1
        Col3 = ""
        Mode3 = 1
        Call TriMultiple("PREPARATIONS", Col1, Mode1, Col2, Mode2, Col3, Mode3, 8, NbPreparations + 1)

        PTotPrepa = 0
        QuantTotPrepa = 0
        If NbPreparations > 0 Then
            For i = 1 To NbPreparations
                Preparation(i) = wsExcel.Cells(i + 1, 1).value
                TestPrepa = wsExcel.Cells(i + 1, 2).value
                TestPrepa = TestPrepa.Substring(0, 1)
                TestPrepa = TestPrepa.ToUpper()
                TaillePrepa(i) = TestPrepa
                Select Case TestPrepa
                    Case "P"            'Taille petite => poids �quivalent viande = 40 gr
                        PoidsPrepa(i) = 10 * CoefPrepa
                    Case "M"            'Taille moyenne => poids = 80 gr
                        PoidsPrepa(i) = 20 * CoefPrepa
                    Case "G"            'Taille grande => poids = 120 gr
                        PoidsPrepa(i) = 30 * CoefPrepa
                    Case Else
                        TexteMsg = "Preparation " & Preparation(i) & " Taille Petit-Moyen-Gros non reconnue"
                        Call Reporting("PREPARATIONS", "ALERTE", TexteMsg, "PREPARATIONS")
                        NbErreur += 1
                End Select
                PoidsTest = wsExcel.Cells(i + 1, 3).value
                If VarType(wsExcel.Cells(i + 1, 3).value) = 5 Then
                    QuantPrepa(i) = wsExcel.Cells(i + 1, 3).value
                    QuantTotPrepa += QuantPrepa(i)
                    PTotPrepa += QuantPrepa(i) * PoidsPrepa(i)
                Else
                    TexteMsg = "Quantit� " & PoidsTest & "  � la ligne " & i + 1 & " n'est pas un nombre!"
                    Call Reporting("PREPARATIONS", "ALERTE", TexteMsg, "PREPARATIONS")
                    NbErreur += 1
                End If

                wsExcel.Cells(i + 1, 8) = PoidsPrepa(i)
                PrepaSC(i) = False
                If wsExcel.Cells(i + 1, 4).value = 1 Then PrepaSC(i) = True
                PrepaSV(i) = False
                If wsExcel.Cells(i + 1, 5).value = 1 Then PrepaSV(i) = True

            Next i
        End If

        '***********************************************************
        '  Lecture des salades
        '***********************************************************

        If FeuilleExiste("SALADES") = False Then
            Call Reporting("SALADES", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("SALADES")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbSalades = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbSalades > MaxDenrees Then
            Call Reporting("SALADES", "ARRET", "Nombre de salades d�passe la dimension > " & MaxDenrees, "SALADES")
            Call Colexit()
            Exit Sub
        End If

        Col1 = "B"
        Mode1 = 1       ' tri croissant
        Col2 = ""
        Mode2 = 1
        Col3 = ""
        Mode3 = 1
        Call TriMultiple("SALADES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 8, NbSalades + 1)

        PTotSalad = 0
        QuantTotSalad = 0
        If NbSalades > 0 Then
            wsExcel.Cells(1, 8).value = "Eqv Poids"
            For i = 1 To NbSalades
                Salade(i) = wsExcel.Cells(i + 1, 1).value
                TestPrepa = wsExcel.Cells(i + 1, 2).value
                TestPrepa = TestPrepa.Substring(0, 1)
                TestPrepa = TestPrepa.ToUpper()
                TailleSalade(i) = TestPrepa
                Select Case TestPrepa
                    Case "P"
                        PoidsSalade(i) = 10 * CoefSalade
                    Case "M"
                        PoidsSalade(i) = 20 * CoefSalade
                    Case "G"
                        PoidsSalade(i) = 30 * CoefSalade
                    Case Else
                        TexteMsg = "Salade " & Salade(i) & " Taille Petit-Moyen-Gros non reconnue"
                        Call Reporting("SALADES", "ALERTE", TexteMsg, "SALADES")
                        NbErreur += 1
                End Select
                PoidsTest = wsExcel.Cells(i + 1, 3).value
                If VarType(wsExcel.Cells(i + 1, 3).value) = 5 Then
                    QuantSalade(i) = wsExcel.Cells(i + 1, 3).value
                    QuantTotSalad += QuantSalade(i)
                    PTotSalad += QuantSalade(i) * PoidsSalade(i)
                Else
                    TexteMsg = "Quantit� " & PoidsTest & "  � la ligne " & i + 1 & " n'est pas un nombre!"
                    Call Reporting("SALADES", "ALERTE", TexteMsg, "SALADES")
                    NbErreur += 1
                End If

                wsExcel.Cells(i + 1, 8) = PoidsSalade(i)
                SaladeSC(i) = False
                If wsExcel.Cells(i + 1, 4).value = 1 Then SaladeSC(i) = True
                SaladeSV(i) = False
                If wsExcel.Cells(i + 1, 5).value = 1 Then SaladeSV(i) = True

            Next i
        End If

        '*****************************************************
        '  lecture des LAITAGES
        '*****************************************************
        If FeuilleExiste("LAITAGES") = False Then
            Call Reporting("LAITAGES", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("LAITAGES")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbLaitages = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbLaitages > MaxDenrees Then
            Call Reporting("LAITAGES", "ARRET", "Nombre de laitages d�passe la dimension > " & MaxDenrees, "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        Col1 = "C"
        Mode1 = 1       ' tri croissant
        Col2 = "D"
        Mode2 = 2
        Col3 = "B"
        Mode3 = 1
        Call TriMultiple("LAITAGES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 6, NbLaitages + 1)

        PtotLait = 0
        PtotZeu = 0

        If NbLaitages > 0 Then
            For i = 1 To NbLaitages

                Laitage(i) = wsExcel.Cells(i + 1, 1).value
                PoidsTest = wsExcel.Cells(i + 1, 2).value
                If VarType(wsExcel.Cells(i + 1, 2).value) = 5 Then
                    QuantLait(i) = wsExcel.Cells(i + 1, 2).value
                Else
                    TexteMsg = "Quantit� " & PoidsTest & "  � la ligne " & i + 1 & " n'est pas un nombre!"
                    Call Reporting("LAITAGES", "ALERTE", TexteMsg, "LAITAGES")
                    NbErreur += 1
                End If

                TestPrepa = wsExcel.Cells(i + 1, 3).value
                TestPrepa = TestPrepa.ToUpper()
                CatLait(i) = TestPrepa
                Equiv(i) = wsExcel.Cells(i + 1, 4).value
                If CatLait(i) = "ZEU" Then
                    PtotZeu += QuantLait(i) * Equiv(i)
                Else
                    PtotLait += QuantLait(i) * Equiv(i)
                End If
                CPLait = wsExcel.Cells(i + 1, 5).value

            Next i
        End If

        '*********************************************************
        '  lecture des FAMILLES
        '*********************************************************
        If FeuilleExiste("FAMILLES") = False Then
            Call Reporting("FAMILLES", "ARRET", "Feuille manquante", "RAPPORT")
            Call Colexit()
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("FAMILLES")
        wsExcel.Activate()

        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbFamilles = wsExcel.Cells(1, 100).value - 1
        wsExcel.Range("A1").Select()

        If NbFamilles > MaxFamilles Then
            Call Reporting("FAMILLES", "ARRET", "Nombre de laitages d�passe la dimension > " & MaxFamilles, "RAPPORT")
            Call Colexit()
            Exit Sub
        End If

        wsExcel.Cells(1, 8) = "Test SCSV"
        wsExcel.Cells(1, 10) = "Random"

        For i = 1 To NbFamilles
            ' concat�nation du test SC et du test SV en colonne 8 (H) (pour effectuer le tri sur 3 colonnes maxi)
            wsExcel.Cells(i + 1, 8).value = wsExcel.Cells(i + 1, 7).value * 10 + wsExcel.Cells(i + 1, 6).value
            ' attribue un ordre al�atoire � chaque famille 
            wsExcel.Cells(i + 1, 10).value = A1.Next(NbFamilles)
            ' calcul le nbre de b�n�ficiaires: une part par adulte et une demi-part par enfant, arrondi � l'unit� sup
            wsExcel.Cells(i + 1, 5).value = Math.Round(wsExcel.Cells(i + 1, 3).value + 0.51 * wsExcel.Cells(i + 1, 4).value)
        Next

        Col1 = "H"
        Mode1 = 2       ' tri croissant
        Col2 = "E"
        Mode2 = 2
        Col3 = "J"
        Mode3 = 1
        Call TriMultiple("FAMILLES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 10, NbFamilles + 1)

        NbTotViande = 0
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).value
            NBenef(i) = wsExcel.Cells(i + 1, 5).value
            SansCochon(i) = False
            SansViande(i) = False
            If wsExcel.Cells(i + 1, 6).value = 1 Then SansCochon(i) = True
            If wsExcel.Cells(i + 1, 7).value = 1 Then SansViande(i) = True
            TestSCSV(i) = wsExcel.Cells(i + 1, 8).value
            NbTotViande += NBenef(i)
            Panier(i) = 0
        Next
        ' teste le nombre de b�n�ficiaires
        If NbTotViande = 0 Then
            Call Reporting("FAMILLES", "ALERTE", " Le nombre de b�n�ficiaires est nul", "FAMILLES")
            NbErreur += 1
        End If

        'calcul du poids th�orique par familles = poids total reparti prorata nbre de b�n�ficiaires
        ModuleViande = PTotViande / NbTotViande
        For i = 1 To NbFamilles
            PoidsTheo(i) = ModuleViande * NBenef(i)
        Next

        '*********************************************************
        ' Mise en forme onglet RESULTATS
        '*********************************************************

        If FeuilleExiste("RESULTATS") = True Then
            appExcel.Application.DisplayAlerts = False
            wbExcel.Worksheets("RESULTATS").delete()        'supprime la feuille avant de commencer
            appExcel.DisplayAlerts = True
        End If
        wsExcel = wbExcel.Worksheets.Add()                  'ajoute une nouvelle feuille
        wsExcel.Name = "RESULTATS"


        wsExcel.Cells(1, 1).value = "N� CAISSE"
        wsExcel.Cells(1, 2).value = "FAMILLE"
        wsExcel.Cells(1, 3).value = "B�n�ficiaires"
        wsExcel.Cells(1, 4).value = "Sans Cochon"
        wsExcel.Cells(1, 5).value = "Sans Viande"

        For i = 1 To NbFamilles
            wsExcel.Cells(i + 1, 1).value = NumFamille(i)
            wsExcel.Cells(i + 1, 2).value = NomFamille(i)
            wsExcel.Cells(i + 1, 3).value = NBenef(i)
            If SansCochon(i) Then wsExcel.Cells(i + 1, 4).value = "OUI"
            If SansViande(i) Then wsExcel.Cells(i + 1, 5).value = "OUI"
        Next

        Decal = 5

        If NbDenrees > 0 Then
            For i = 1 To NbDenrees
                wsExcel.Cells(1, i + Decal).value = Description(i) & " " & Poids(i) & " Gr (" & Quant(i) & ")"
            Next
            wsExcel.Cells(1, NbDenrees + 6).value = "Poids attribu�"
            wsExcel.Cells(1, NbDenrees + 7).value = "Poids th�orique"
            wsExcel.Cells(1, NbDenrees + 8).value = "Ecart"
            wsExcel.Cells(NbFamilles + 2, 1).value = "SOMME"
            wsExcel.Cells(NbFamilles + 2, 2).value = "SOMME"


            '***************************************************************
            '   ATTRIBUTION DES VIANDES
            '***************************************************************
            'priorit� 1: familles sans viande et sans cochon

            Call Attribution1(NbDenrees, ResteQuant, Poids, ViandeSC, ViandeSV)

            Call Attribution2(NbDenrees, ModuleViande, ResteQuant, Poids, ViandeSC, ViandeSV)

            ParamEcart = 1             'type d'�cart = �cart calcul�
            Call Attribution3(NbDenrees, ResteQuant, Poids, ViandeSC, ViandeSV, ParamEcart)


            'impression des r�sultats

            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, NbDenrees + 6).value = Panier(i)
                wsExcel.Cells(i + 1, NbDenrees + 7).value = PoidsTheo(i)
                wsExcel.Cells(i + 1, NbDenrees + 8).value = Panier(i) - PoidsTheo(i)
            Next

            For j = 1 To NbDenrees
                wsExcel.Cells(NbFamilles + 2, j + Decal).value = "=SUM(R[-" & NbFamilles & "]C:R[-1]C)"
            Next

            wsExcel.Cells(NbFamilles + 2, NbDenrees + 7).value = PTotViande

            wsExcel = wbExcel.Worksheets("FAMILLES")
            wsExcel.Cells(1, 9) = "ECART"
            ' reporte l'�cart entre la dotation th�orique et r�alis� pour prioriser l'attribution des plats pr�par�s
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, 9) = Panier(i) - PoidsTheo(i)
            Next

        End If      'Fin du test s'il n'y a pas de viande

        '*******************************************************************************
        '   PLATS PREPARES
        '*******************************************************************************
        ' le tri des familles et RESULTATS se fait m�me sans plats prepar�s car 
        ' il sert aussi aux salades
        If NbDenrees > 0 Then Decal = Decal + NbDenrees + 3

        'tri des familles
        wsExcel = wbExcel.Worksheets("FAMILLES")
        Col1 = "G"
        Mode1 = 2       ' tri descending
        Col2 = "F"
        Mode2 = 2
        If NbDenrees > 0 Then
            Col3 = "I"      ' tri sur les �carts de poids (attribu� - th�orique) pour prioriser l'attribution des pr�parations
            Mode3 = 1
        Else
            Col3 = "E"
            Mode3 = 2
        End If
        Call TriMultiple("FAMILLES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 10, NbFamilles + 1)

        'on relit les familles apr�s le tri
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).value
            NBenef(i) = wsExcel.Cells(i + 1, 5).value
            SansCochon(i) = False
            SansViande(i) = False
            If wsExcel.Cells(i + 1, 6).value = 1 Then SansCochon(i) = True
            If wsExcel.Cells(i + 1, 7).value = 1 Then SansViande(i) = True
            TestSCSV(i) = wsExcel.Cells(i + 1, 8).value
            Panier(i) = wsExcel.Cells(i + 1, 9).value
        Next

        ' tri des r�sultats, de la m�me fa�on
        wsExcel = wbExcel.Worksheets("RESULTATS")
        Col1 = "E"
        Mode1 = 2       ' tri descending
        Col2 = "D"
        Mode2 = 2
        If NbDenrees > 0 Then
            Col3 = AlphaCol(Decal)
            Mode3 = 1
        Else
            Col3 = "C"
            Mode3 = 2
        End If
        k = NbDenrees + NbPreparations + NbSalades + NbLaitages + 30
        Call TriMultiple("RESULTATS", Col1, Mode1, Col2, Mode2, Col3, Mode3, k, NbFamilles + 2)

        If NbPreparations > 0 Then
            wsExcel = wbExcel.Worksheets("RESULTATS")
            'initialisation du reste � la quantit� initiale
            For i = 1 To NbPreparations
                ResteQuant(i) = QuantPrepa(i)
            Next i
            'calcul du poids th�orique par famille
            For i = 1 To NbFamilles
                PoidsTheo(i) = PTotPrepa * NBenef(i) / NbTotViande
            Next
            'en t�te de l'onglet RESULTATS
            For i = 1 To NbPreparations
                wsExcel.Cells(1, i + Decal).value = Preparation(i) & " " & TaillePrepa(i) & " (" & QuantPrepa(i) & ")"
            Next
            wsExcel.Cells(1, NbPreparations + Decal + 1).value = "Nbre attribu�"

            '*******************************************************************
            ' Attribution des Plats Prepar�s
            '*******************************************************************
            Call Attribution4(NbPreparations, ResteQuant, PoidsPrepa, PrepaSC, PrepaSV)

            ParamEcart = 2      'ecart donn� par l'ordre des familles, en colonne I
            Call Attribution3(NbPreparations, ResteQuant, PoidsPrepa, PrepaSC, PrepaSV, ParamEcart)

            'total de chaque plat en bas de tableau 
            For j = 1 To NbPreparations
                wsExcel.Cells(NbFamilles + 2, j + Decal).value = "=SUM(R[-" & NbFamilles & "]C:R[-1]C)"
            Next
            'report nbre de plats prepar�s par famille
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, Decal + NbPreparations + 1).value = "=SUM(RC[-" & NbPreparations & "]:RC[-1])"
            Next

            wsExcel = wbExcel.Worksheets("FAMILLES")
            wsExcel.Cells(1, 9).value = "Ecart"
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, 9).value = Panier(i) - PoidsTheo(i)
            Next
        End If

        '************************************************************************
        '    SALADES
        '************************************************************************

        If NbPreparations > 0 Then Decal = Decal + NbPreparations + 1

        If NbSalades > 0 Then

            wsExcel = wbExcel.Worksheets("FAMILLES")
            For i = 1 To NbFamilles
                Panier(i) = wsExcel.Cells(i + 1, 7).value
                PoidsTheo(i) = PTotSalad * NBenef(i) / NbTotViande
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")
            QuantTotSalad = 0

            For i = 1 To NbSalades
                ResteQuant(i) = QuantSalade(i)
                wsExcel.Cells(1, i + Decal).value = Salade(i) & " " & TailleSalade(i) & " (" & QuantSalade(i) & ")"
                QuantTotSalad += QuantSalade(i)
            Next
            wsExcel.Cells(1, Decal + NbSalades + 1).value = "Nbre Attribu�"

            '***********************************************************************
            ' Attribution des SALADES
            '***********************************************************************
            Call Attribution4(NbSalades, ResteQuant, PoidsSalade, SaladeSC, SaladeSV)

            ParamEcart = 2      'ecart donn� par l'ordre des familles, en colonne I
            Call Attribution3(NbSalades, ResteQuant, PoidsSalade, SaladeSC, SaladeSV, ParamEcart)

            For j = 1 To NbSalades
                wsExcel.Cells(NbFamilles + 2, j + Decal).value = "=SUM(R[-" & NbFamilles & "]C:R[-1]C)"
            Next

            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, Decal + NbSalades + 1).value = "=SUM(RC[-" & NbSalades & "]:RC[-1])"
            Next

        End If

        '*****************************************************
        '  LAITAGES
        '*****************************************************

        'tri des familles
        wsExcel = wbExcel.Worksheets("FAMILLES")
        Col1 = "E"
        Mode1 = 2       ' tri descending
        Col2 = "A"
        Mode2 = 1
        Col3 = ""
        Mode3 = 1

        Call TriMultiple("FAMILLES", Col1, Mode1, Col2, Mode2, Col3, Mode3, 10, NbFamilles + 1)

        'on relit les familles apr�s le tri
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).value
            NBenef(i) = wsExcel.Cells(i + 1, 5).value

            Panier(i) = 0
            PanierZeu(i) = 0
        Next

        wsExcel.Range("H1:J300").Clear()

        ' tri des r�sultats, de la m�me fa�on
        wsExcel = wbExcel.Worksheets("RESULTATS")
        Col1 = "C"
        Mode1 = 2       ' tri descending
        Col2 = "A"
        Mode2 = 1
        Col3 = ""
        Mode3 = 2

        k = NbDenrees + NbPreparations + NbSalades + NbLaitages + 30

        Call TriMultiple("RESULTATS", Col1, Mode1, Col2, Mode2, Col3, Mode3, k, NbFamilles + 2)

        If NbLaitages > 0 Then
            PtotLait = 0
            PtotZeu = 0
            For j = 1 To NbLaitages
                If CatLait(j) = "ZEU" Then
                    PtotZeu += QuantLait(j) * Equiv(j)
                Else
                    If CatLait(j) <> "BCF" Then PtotLait += QuantLait(j) * Equiv(j)
                End If
                ResteQuant(j) = QuantLait(j)
            Next

            For i = 1 To NbFamilles
                PoidsTheo(i) = PtotLait * NBenef(i) / NbTotViande
                PoidsTheozeu(i) = PtotZeu * NBenef(i) / NbTotViande
            Next

            If NbSalades > 0 Then Decal += NbSalades + 1

            For i = 1 To NbLaitages
                wsExcel.Cells(1, i + Decal).value = Laitage(i) & "_EQV" & Equiv(i) & " (" & QuantLait(i) & ")"
            Next
            wsExcel.Cells(1, NbLaitages + Decal + 1).value = "Nbre Laitages"
            wsExcel.Cells(1, NbLaitages + Decal + 2).value = "Nbre d'oeufs"

            '********************************************************************
            ' repartition des laitages
            '********************************************************************
            '---remplissage panier sous limite, sauf equiv 0,5, BCF et Zeu---------------
            For i = 1 To NbFamilles
                If NBenef(i) > 0 Then
                    For j = 1 To NbLaitages
                        CibleLait = QuantLait(j) * NBenef(i) / NbTotViande
                        If CatLait(j) <> "BCF" And CatLait(j) <> "ZEU" Then
                            If Equiv(j) > 1 Then
                                If ResteQuant(j) > 0 Then
                                    TestPanier = Panier(i) + Equiv(j)
                                    If TestPanier < PoidsTheo(i) Then
                                        wsExcel.Cells(i + 1, Decal + j).value += 1
                                        Panier(i) += Equiv(j)
                                        ResteQuant(j) -= 1
                                    End If
                                End If
                            Else
                                If Equiv(j) = 1 Then
                                    If ResteQuant(j) > 2 Then
                                        TestPanier = 2
                                        Do While TestPanier < CibleLait
                                            wsExcel.Cells(i + 1, Decal + j).value += 2
                                            Panier(i) += 2
                                            TestPanier += 2
                                            ResteQuant(j) -= 2
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
            Next i


            '------quantit�s restantes, on donne un laitage � chacun ------------
            For i = 1 To NbFamilles
                Saut = False
                If Panier(i) = 0 And NBenef(i) > 0 Then
                    For j = 1 To NbLaitages
                        If Equiv(j) >= 1 And ResteQuant(j) > 0 And Saut = False _
                            And CatLait(j) <> "BCF" And CatLait(j) <> "ZEU" Then
                            wsExcel.Cells(i + 1, Decal + j).value += 1
                            Panier(i) += Equiv(j)
                            ResteQuant(j) -= 1
                            Saut = True
                        End If
                    Next j
                End If
            Next i

            '-----vide le stock 2 par 2 -----------------------------------------
            For j = 1 To NbLaitages
                If ResteQuant(j) > 1 And CatLait(j) <> "BCF" And CatLait(j) <> "ZEU" Then
                    While ResteQuant(j) > 1
                        EcartMaxi = 0
                        NumeMaxi = 1
                        For i = 1 To NbFamilles
                            Ecart = Panier(i) - PoidsTheo(i)
                            If Ecart < EcartMaxi And NBenef(i) > 0 Then
                                EcartMaxi = Ecart
                                NumeMaxi = i
                            End If
                        Next i
                        wsExcel.Cells(NumeMaxi + 1, Decal + j).value += 2
                        Panier(NumeMaxi) += Equiv(j) * 2
                        ResteQuant(j) -= 2
                    End While
                End If
            Next j

            '--------repartition des impairs------------------------------
            For j = 1 To NbLaitages
                If ResteQuant(j) > 1 And CatLait(j) <> "BCF" And CatLait(j) <> "ZEU" Then
                    Do While ResteQuant(j) > 0
                        Saut = False
                        For i = 1 To NbFamilles
                            Ecart = Panier(i) - PoidsTheo(i)
                            If Ecart < 0 And Saut = False And NBenef(i) > 0 Then
                                wsExcel.Cells(NumeMaxi + 1, Decal + j).value += 1
                                Panier(NumeMaxi) += Equiv(j)
                                ResteQuant(j) -= 1
                                Saut = True
                            End If
                        Next
                    Loop

                End If
            Next j

            '-----repartition des ZEU---------------------------------------
            For i = 1 To NbFamilles
                If NBenef(i) > 0 Then
                    Saut = False
                    For j = 1 To NbLaitages
                        If CatLait(j) = "ZEU" Then
                            For k = 1 To ResteQuant(j)
                                If Saut = False Then
                                    TestPanier = PanierZeu(i) + Equiv(j)
                                    If TestPanier < PoidsTheozeu(i) Then
                                        wsExcel.Cells(i + 1, Decal + j).value += 1
                                        PanierZeu(i) += Equiv(j)
                                        ResteQuant(j) -= 1
                                    Else
                                        If Equiv(j) > PoidsTheozeu(i) Then
                                            wsExcel.Cells(i + 1, Decal + j).value += 1
                                            PanierZeu(i) += Equiv(j)
                                            ResteQuant(j) -= 1
                                            Saut = True
                                        End If
                                    End If
                                End If
                            Next k
                        End If
                    Next j
                End If
            Next i

            '-----Vide la stock BCF ou ZEU ----------------------------------
            i = 0
            For j = 1 To NbLaitages
                If CatLait(j) = "BCF" Or CatLait(j) = "ZEU" Then
                    Do While ResteQuant(j) > 0
                        i += 1
                        If i > NbFamilles Then i = 1
                        If NBenef(i) > 0 Then
                            wsExcel.Cells(i + 1, Decal + j).value += 1
                            If CatLait(j) = "ZEU" Then PanierZeu(i) += Equiv(j)
                            ResteQuant(j) -= 1
                        End If
                    Loop
                End If
            Next

            '---ecriture RESULTATS----------------------------------------------
            For i = 1 To NbFamilles
                SommeLait = 0
                SommeZeu = 0
                For j = 1 To NbLaitages
                    If CatLait(j) = "ZEU" Then
                        SommeZeu += wsExcel.Cells(i + 1, Decal + j).value
                    Else
                        SommeLait += wsExcel.Cells(i + 1, Decal + j).value
                    End If
                Next
                wsExcel.Cells(i + 1, Decal + NbLaitages + 1).value = SommeLait
                wsExcel.Cells(i + 1, Decal + NbLaitages + 2).value = PanierZeu(i)
            Next
            For j = 1 To NbLaitages
                wsExcel.Cells(NbFamilles + 2, j + Decal).value = "=SUM(R[-" & NbFamilles & "]C:R[-1]C)"
            Next
            Decal += NbLaitages + 2
        End If

        '*********************************************************************
        'Formatage onglet DIVERS
        '**********************************************************************
        If FeuilleExiste("DIVERS") = False Then
            wsExcel = wbExcel.Worksheets.Add()
            wsExcel.Name = "DIVERS"
            NbDivers = 0
        Else
            wsExcel = wbExcel.Worksheets("DIVERS")
            wsExcel.Activate()
            wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
            NbDivers = wsExcel.Cells(1, 100).value - 1
            wsExcel.Range("A1").Select()
        End If

        If NbDivers > 0 Then
            For i = 1 To NbDivers
                Divers(i) = wsExcel.Cells(i + 1, 1).value
            Next
            wsExcel = wbExcel.Worksheets("RESULTATS")
            For i = 1 To NbDivers
                wsExcel.Cells(1, i + Decal).value = Divers(i)
                wsExcel.Cells(NbFamilles + 2, i + Decal).value = "=SUM(R[-" & NbFamilles & "]C:R[-1]C)"
            Next i
        End If

        '**********************************************************************
        '  formattage onglet RESULTATS
        '**********************************************************************

        wsExcel = wbExcel.Worksheets("RESULTATS")
        'mise en gras colonnes familles
        wsExcel.Columns("A:E").Font.Bold = True
        wsExcel.Columns("A").HorizontalAlignment = -4108
        wsExcel.Columns("C:E").ColumnWidth = 4
        wsExcel.Columns("C:E").HorizontalAlignment = -4108

        'Premi�re ligne
        Decal = 5
        If NbDenrees > 0 Then Decal += NbDenrees + 3
        If NbPreparations > 0 Then Decal += NbPreparations + 1
        If NbSalades > 0 Then Decal += NbSalades + 1
        If NbLaitages > 0 Then Decal += NbLaitages + 2
        If NbDivers > 0 Then Decal += NbDivers

        AlphaColTri = AlphaCol(Decal)
        With wsExcel.Range("C1:" & AlphaColTri & "1")
            .HorizontalAlignment = -4108
            .Orientation = 90
            .Font.Bold = True
            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        End With

        i = 0
        While i < NbFamilles
            i += 2
            wsExcel.Range("A" & i & ":" & AlphaColTri & i).Interior.ColorIndex = 19
        End While

        Decal = 5
        If NbDenrees > 0 Then
            AlphaColTri = AlphaCol(Decal + NbDenrees)
            wsExcel.Columns("F:" & AlphaColTri).ColumnWidth = 4
            wsExcel.Columns("F:" & AlphaColTri).HorizontalAlignment = -4108
            Call TraitsVerticaux("A", AlphaColTri, NbFamilles + 2)

            AlphaColTri = AlphaCol(Decal + NbDenrees + 1)
            AlphaColTri2 = AlphaCol(Decal + NbDenrees + 3)
            With wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2)
                .ColumnWidth = 6
                .font.bold = True
                .numberformat = "0.0"
            End With
            wsExcel.Range(AlphaColTri & 1 & ":" & AlphaColTri2 & NbFamilles + 2).Interior.ColorIndex = 40

            Call Bordure(AlphaColTri, AlphaColTri2, NbFamilles + 2)
            Decal += NbDenrees + 3

        End If
        If NbPreparations > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbPreparations)
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).ColumnWidth = 4
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).HorizontalAlignment = -4108
            Call TraitsVerticaux(AlphaColTri, AlphaColTri2, NbFamilles + 2)

            AlphaColTri = AlphaCol(Decal + NbPreparations + 1)
            With wsExcel.Columns(AlphaColTri & ":" & AlphaColTri)
                .ColumnWidth = 6
                .font.bold = True
                .HorizontalAlignment = -4108
            End With
            wsExcel.Range(AlphaColTri & 1 & ":" & AlphaColTri & NbFamilles + 2).Interior.ColorIndex = 40

            Call Bordure(AlphaColTri, AlphaColTri, NbFamilles + 2)
            Decal += NbPreparations + 1
        End If
        If NbSalades > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbSalades)
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).ColumnWidth = 4
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).HorizontalAlignment = -4108
            Call TraitsVerticaux(AlphaColTri, AlphaColTri2, NbFamilles + 2)

            AlphaColTri = AlphaCol(Decal + NbSalades + 1)
            With wsExcel.Columns(AlphaColTri & ":" & AlphaColTri)
                .ColumnWidth = 6
                .font.bold = True
                .HorizontalAlignment = -4108
            End With
            wsExcel.Range(AlphaColTri & 1 & ":" & AlphaColTri & NbFamilles + 2).Interior.ColorIndex = 40

            Call Bordure(AlphaColTri, AlphaColTri, NbFamilles + 2)
            Decal += NbSalades + 1
        End If
        If NbLaitages > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbLaitages)
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).ColumnWidth = 4
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).HorizontalAlignment = -4108
            Call TraitsVerticaux(AlphaColTri, AlphaColTri2, NbFamilles + 2)

            AlphaColTri = AlphaCol(Decal + NbLaitages + 1)
            AlphaColTri2 = AlphaCol(Decal + NbLaitages + 2)
            With wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2)
                .ColumnWidth = 6
                .font.bold = True
                .HorizontalAlignment = -4108
            End With
            wsExcel.Range(AlphaColTri & 1 & ":" & AlphaColTri2 & NbFamilles + 2).Interior.ColorIndex = 40
            Call Bordure(AlphaColTri, AlphaColTri, NbFamilles + 2)
            Decal += NbLaitages + 1
        End If
        If NbDivers > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbDivers + 1)
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).ColumnWidth = 4
            wsExcel.Columns(AlphaColTri & ":" & AlphaColTri2).HorizontalAlignment = -4108
            Call TraitsVerticaux(AlphaColTri, AlphaColTri2, NbFamilles + 2)
        End If

        Call Colexit()

    End Sub

    Private Sub Attribution1(Nbdenrees As Integer, Reste() As Integer, Poids() As Single, Testsc() As Boolean, Testsv() As Boolean)
        'attribution prioritaire aux familles SansViande et Sanscochon
        '-------------------------------------------------------------
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim TestPanier As Single
        Dim Saut As Boolean

        For i = 1 To NbFamilles
            If SansViande(i) And NBenef(i) > 0 Then
                Saut = False
                For j = 1 To Nbdenrees
                    If (SansViande(i) = False Or Testsv(j) = True) And
                        (SansCochon(i) = False Or Testsc(j) = True) And Saut = False Then
                        For k = 1 To Reste(j)
                            TestPanier = Reste(j) + Panier(i)
                            If TestPanier < PoidsTheo(i) Then
                                wsExcel.Cells(i + 1, Decal + j).value += 1
                                Panier(i) += Poids(j)
                                Reste(j) -= 1
                            Else
                                If Poids(j) > PoidsTheo(i) Then
                                    wsExcel.Cells(i + 1, Decal + j).value += 1
                                    Panier(i) += Poids(j)
                                    Reste(j) -= 1
                                    Saut = True
                                    Exit For
                                End If
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i

    End Sub

    Private Sub Attribution2(Nbdenrees As Integer, ModuleViande As Single, Reste() As Integer, Poids() As Single, Testsc() As Boolean, Testsv() As Boolean)
        ' attribution des denrees les plus lourdes: un exemplaire par famille
        '--------------------------------------------------------------------
        Dim i As Integer
        Dim j As Integer

        i = 0
        For j = 1 To Nbdenrees
            If Poids(j) > ModuleViande Then
                Do While Reste(j) > 0
                    i += 1
                    If i > NbFamilles Then i = 1
                    If (SansViande(i) = False Or Testsv(j) = True) And
                       (SansCochon(i) = False Or Testsc(j) = True) And NBenef(i) > 0 Then
                        wsExcel.Cells(i + 1, Decal + j).value += 1
                        Panier(i) += Poids(j)
                        Reste(j) -= 1
                    End If
                Loop
            End If
        Next j
    End Sub

    Private Sub Attribution3(Nbdenrees As Integer, Reste() As Integer, Poids() As Single, Testsc() As Boolean,
                             Testsv() As Boolean, ParamEcart As Integer)
        'attribution des denr�es: vider les stock denr�e par denr�e
        '-----------------------------------------------------------
        Dim i As Integer
        Dim j As Integer
        'Dim k As Integer
        Dim NumeMaxi As Integer
        Dim EcartMaxi As Single
        Dim Ecart As Single

        NumeMaxi = 0
        For j = 1 To Nbdenrees
            Do While Reste(j) > 0
                If ParamEcart = 1 Then                  'calcul de l'ecart maxi
                    EcartMaxi = PoidsTheo(1)
                    NumeMaxi = 0
                    For i = 1 To NbFamilles             'calcul de la famille en �cart maxi
                        If (SansViande(i) = False Or Testsv(j) = True) And
                            (SansCochon(i) = False Or Testsc(j) = True) And NBenef(i) > 0 Then
                            Ecart = Panier(i) - PoidsTheo(i)
                            If Ecart < EcartMaxi Then
                                EcartMaxi = Ecart
                                NumeMaxi = i
                            End If
                        End If
                    Next i

                    If NumeMaxi = 0 Then
                        ' pas de famille trouv�e en �cart maxi   = probl�me
                        TexteMsg = "(Sub. Attribution3) Arr�t � la denr�e " & "j " & " sur " & Nbdenrees & " R�partition incompl�te"
                        Call Reporting("RESULTATS", "ARRET", TexteMsg, "RESULTATS")
                        Call Colexit()
                        Exit Sub
                    End If
                Else
                    NumeMaxi += 1
                    If NumeMaxi > NbFamilles Then NumeMaxi = 1
                End If
                'attribution des denr�es une par une
                If (SansViande(NumeMaxi) = False Or Testsv(j) = True) And
                    (SansCochon(NumeMaxi) = False Or Testsc(j) = True) And NBenef(NumeMaxi) > 0 Then
                    wsExcel.Cells(NumeMaxi + 1, Decal + j).value += 1
                    Panier(NumeMaxi) += Poids(j)
                    Reste(j) -= 1
                End If
            Loop
        Next j

    End Sub

    Private Sub Attribution4(Nbdenrees As Integer, Reste() As Integer, Poids() As Single, Testsc() As Boolean, Testsv() As Boolean)
        'par famille, remplissage panier sous la limite
        '----------------------------------------------------------------
        Dim i As Integer
        Dim j As Integer
        Dim TestPanier As Single

        For i = 1 To NbFamilles
            If NBenef(i) > 0 Then
                For j = 1 To Nbdenrees
                    If (SansViande(i) = False Or Testsv(j) = True) And
                        (SansCochon(i) = False Or Testsc(j) = True) And Reste(j) > 0 Then
                        TestPanier = Poids(j) + Panier(i)
                        If TestPanier < PoidsTheo(i) Then
                            wsExcel.Cells(i + 1, Decal + j).value += 1
                            Panier(i) += Poids(j)
                            Reste(j) -= 1
                        End If
                    End If
                Next j
            End If
        Next i

    End Sub

    Public Function FeuilleExiste(FeuilleAVerifier As String) As Boolean
        'fonction qui v�rifie si la "FeuilleAVerifier" existe dans le Classeur actif

        On Error GoTo SiErreur
        Dim Feuille
        wbExcel.Activate()
        FeuilleExiste = False
        For Each Feuille In wbExcel.Sheets
            If UCase(Feuille.Name) = UCase(FeuilleAVerifier) Then
                FeuilleExiste = True
                Exit Function
            End If
        Next Feuille
        Exit Function

SiErreur:
        FeuilleExiste = "Erreur"
    End Function

    Public Function AlphaCol(k As Integer) As String
        ' fonction de conversion du num�ro de colonne en lettre
        Dim h As Integer
        h = Fix((k - 1) / 26)
        If h > 0 Then AlphaCol = Chr(64 + h) & Chr(64 + k - (26 * h)) Else AlphaCol = Chr(64 + k)

    End Function
    Public Sub TriMultiple(Feuille As String, Col1 As String, Mode1 As Integer, Col2 As String, Mode2 As Integer,
                    Col3 As String, Mode3 As Integer, nbcol As Integer, nblignes As Integer)

        Dim h As Integer
        Dim AlphaColTri As String

        wsExcel = wbExcel.Worksheets(Feuille)
        wsExcel.Activate()
        h = nbcol
        AlphaColTri = AlphaCol(h)

        appExcel.CutCopyMode = False
        wsExcel.Sort.SortFields.Clear()
        wsExcel.Sort.SortFields.Add(Key:=wsExcel.Range(Col1 & "2:" & Col1 & nblignes), Order:=Mode1)
        If Col2 <> "" Then wsExcel.Sort.SortFields.Add(Key:=wsExcel.Range(Col2 & "2:" & Col2 & nblignes), Order:=Mode2)
        If Col3 <> "" Then wsExcel.Sort.SortFields.Add(Key:=wsExcel.Range(Col3 & "2:" & Col3 & nblignes), Order:=Mode3)

        With wsExcel.Sort
            .SetRange(wsExcel.Range("A1:" & AlphaColTri & nblignes))
            .Header = XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = Constants.xlTopToBottom
            .SortMethod = XlSortMethod.xlPinYin
            .Apply()
        End With
        wsExcel.Range("A1").Select()
    End Sub

    Public Sub TraitsVerticaux(Alpha1 As String, Alpha2 As String, Nblignes As Integer)
        With wsExcel.Range(Alpha1 & "1:" & Alpha2 & Nblignes).Borders(XlBordersIndex.xlInsideVertical)
            .LineStyle = XlLineStyle.xlContinuous
            .Weight = XlBorderWeight.xlThin
        End With
    End Sub

    Public Sub Bordure(Alpha1 As String, Alpha2 As String, Nblignes As Integer)
        With wsExcel.Range(Alpha1 & "1:" & Alpha2 & Nblignes)
            .Font.Bold = True
            .Borders(XlBordersIndex.xlEdgeLeft).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Borders(XlBordersIndex.xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
        End With
    End Sub
    Public Sub Reporting(Onglet As String, Criticite As String, ReportMsg As String, Retour As String)

        wsExcel = wbExcel.Worksheets("RAPPORT")
        nbReport += 1
        wsExcel.Cells(nbReport, 1).value = Onglet
        wsExcel.Cells(nbReport, 2).value = ReportMsg
        wsExcel.Cells(nbReport, 8).value = Criticite
        wsExcel = wbExcel.Worksheets("" & Retour & "")
    End Sub

    Public Sub MAJ()
        '*********************************************
        '  MISE A JOUR
        '*********************************************
        Dim NbDenrees As Integer
        Dim NbPreparations As Integer
        Dim NbSalades As Integer
        Dim NbLaitages As Integer
        Dim Poids(MaxDenrees) As Single
        Dim CatLait(MaxDenrees) As String
        Dim Equiv(MaxDenrees) As Single
        Dim i, j As Integer
        Dim SommeZeu As Single
        Dim SommeLait As Single
        Dim Cellule As String
        Dim NbErreur As Integer


        wsExcel = wbExcel.Worksheets("FAMILLES")
        wsExcel.Activate()
        NbFamilles = wsExcel.Cells(1, 100).value - 1

        wsExcel = wbExcel.Worksheets("VIANDES")
        wsExcel.Activate()
        NbDenrees = wsExcel.Cells(1, 100).value - 1

        NbErreur = 0
        Decal = 5
        If NbDenrees > 0 Then
            For j = 1 To NbDenrees
                Poids(j) = wsExcel.Cells(j + 1, 2).value
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")

            For i = 1 To NbFamilles
                For j = 1 To NbDenrees
                    Cellule = wsExcel.Cells(i + 1, Decal + j).value
                    If String.IsNullOrEmpty(Cellule) = False Then
                        If IsNumeric(Cellule) Then
                            Panier(i) += Cellule * Poids(j)
                        Else
                            TexteMsg = "Ligne " & i + 1 & " Col " & j + Decal & "  cellule " & Cellule & "  n'est pas un nombre"
                            Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RESULTATS")
                            NbErreur += 1
                        End If
                    End If
                Next
                wsExcel.Cells(i + 1, NbDenrees + 6).value = Panier(i)
                wsExcel.Cells(i + 1, NbDenrees + 8).value = Panier(i) - wsExcel.Cells(i + 1, NbDenrees + 7).value
            Next
            If NbErreur > 0 Then
                Call Colexit()
                Exit Sub
            End If
            Decal += NbDenrees + 3

        End If

        wsExcel = wbExcel.Worksheets("PREPARATIONS")
        NbPreparations = wsExcel.Cells(1, 100).value - 1
        If NbPreparations > 0 Then Decal += NbPreparations + 1

        wsExcel = wbExcel.Worksheets("SALADES")
        NbSalades = wsExcel.Cells(1, 100).value - 1
        If NbSalades > 0 Then Decal += NbSalades + 1

        wsExcel = wbExcel.Worksheets("LAITAGES")
        NbLaitages = wsExcel.Cells(1, 100).value - 1
        For i = 1 To NbLaitages
            CatLait(i) = (wsExcel.Cells(i + 1, 3).value).toupper()
            Equiv(i) = wsExcel.Cells(i + 1, 4).value
        Next

        wsExcel = wbExcel.Worksheets("RESULTATS")
        NbErreur = 0
        For i = 1 To NbFamilles
            SommeLait = 0
            SommeZeu = 0
            For j = 1 To NbLaitages
                Cellule = wsExcel.Cells(i + 1, Decal + j).value
                If String.IsNullOrEmpty(Cellule) = False Then
                    If IsNumeric(Cellule) Then
                        If CatLait(j) = "ZEU" Then
                            SommeZeu += Cellule * Equiv(j)
                        Else
                            SommeLait += Cellule
                        End If
                    Else
                        wsExcel = wbExcel.Worksheets("RAPPORT")
                        nbReport += 1
                        wsExcel.Cells(nbReport, 1).value = "RESULTATS"
                        wsExcel.Cells(nbReport, 2).value = "Ligne " & i + 1 & " Col " & j + Decal & "  cellule " & Cellule & "  n'est pas un nombre"
                        wsExcel.Cells(nbReport, 8).value = "ALERTE"
                        wsExcel = wbExcel.Worksheets("RESULTATS")
                        NbErreur += 1
                    End If
                End If
            Next
            wsExcel.Cells(i + 1, Decal + NbLaitages + 1).value = SommeLait
            wsExcel.Cells(i + 1, Decal + NbLaitages + 2).value = SommeZeu
        Next

        Call Colexit()

    End Sub

    Public Sub AIDA()

        '*********************************************************
        '   AIDA
        '*********************************************************
        '
        ' Ecriture des r�sultats au format AIDA
        '
        Dim nbdenrees As Integer
        Dim NbPreparations As Integer
        Dim NbSalades As Integer
        Dim NbLaitages As Integer
        Dim NbDivers As Integer
        Dim nbPrix As Integer
        '********* variables des onglets de denr�es*******************
        Dim Poids(MaxDenrees) As Single           ' poids cumul pour Fromages et divers
        Dim CodePrixDenree(MaxDenrees) As String   ' code prix des denr�es
        Dim Equiv(MaxDenrees) As Single           ' equiv yaourt pour les boites � oeufs

        '*********variables de la base PRIX ***********
        Dim CodePrix(MaxPrix) As String
        Dim CodeAIDA(MaxPrix) As String
        Dim UnitAIDA(MaxPrix) As String
        Dim PrixAIDA(MaxPrix) As Single
        '********variables de la liste unique des cat�gories AIDA *********************
        Dim Categorie(MaxPrix) As String         ' code prix de la cat�gorie
        Dim CatAIDA(MaxPrix) As String           ' cat�gorie g�n�rique
        Dim UnAIDA(MaxPrix) As String            ' unit�
        Dim PrixListe(MaxPrix) As Single
        Dim PoidsCat(MaxPrix) As Single          'variable de travail
        Dim PrixPanier As Single                 ' Prix du panier *** dimension du nombre de familles****

        Dim Test As Boolean
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim jdec As Integer
        Dim AlphaColTri As String

        Dim Contenu As String
        Dim Erreur As Boolean
        Dim Ecart As Single
        Dim SousTotal As Single
        Dim TotalArrondi As Single
        Dim Arrondi As Single
        Dim FormatCol As String

        ' chemin r�seau g�n�rique pour l'enregistrement des images de code-barre
        CheminBureau = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

        wsExcel = wbExcel.Worksheets("FAMILLES")
        wsExcel.Activate()
        NbFamilles = wsExcel.Cells(1, 100).value - 1

        wsExcel = wbExcel.Worksheets("PRIX")
        wsExcel.Activate()
        nbPrix = wsExcel.Cells(1, 100).value - 1

        '----------------------lecture des codes prix ---------------------------------------
        If nbPrix > MaxPrix Then
            TexteMsg = "Nombre de prix d�passe la dimension > " & MaxPrix
            Call Reporting("PRIX", "ARRET", TexteMsg, "PRIX")
            Call Colexit()
            Exit Sub
        End If

        If nbPrix > 0 Then
            For i = 1 To nbPrix
                CodePrix(i) = wsExcel.Cells(i + 1, 2).value
                CodeAIDA(i) = wsExcel.Cells(i + 1, 3).value
                UnitAIDA(i) = wsExcel.Cells(i + 1, 4).value
                PrixAIDA(i) = wsExcel.Cells(i + 1, 5).value
                Select Case UCase(UnitAIDA(i))
                    Case "KGM"
                        UnitAIDA(i) = "[KgM]"
                    Case "KGC"
                        UnitAIDA(i) = "[KgC]"
                    Case "BOI"
                        UnitAIDA(i) = "[BOI]"
                    Case "UN"
                        UnitAIDA(i) = "[Un]"
                    Case Else
                End Select
            Next i
        Else
            TexteMsg = "Pas de code prix document�"
            Call Reporting("PRIX", "ARRET", TexteMsg, "PRIX")
            Call Colexit()
        End If

        NbCat = 0

        ' **** CONSTRUCTION DE LA LISTE UNIQUE et SANS DOUBLONS DES CODES PRIX *****
        '-----------------liste des codes viandes------------------------------------
        wsExcel = wbExcel.Worksheets("VIANDES")
        wsExcel.Activate()
        nbdenrees = wsExcel.Cells(1, 100).value - 1

        If nbdenrees > 0 Then
            For j = 1 To nbdenrees
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
            Next

            Call ListeCategorie(nbdenrees, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
        UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)

        End If

        '------------liste des codes plats prepares---------------------------------------
        wsExcel = wbExcel.Worksheets("PREPARATIONS")
        wsExcel.Activate()
        NbPreparations = wsExcel.Cells(1, 100).value - 1

        If NbPreparations > 0 Then
            For j = 1 To NbPreparations
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
            Next j

            Call ListeCategorie(NbPreparations, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes plats prepares---------------------------------------
        wsExcel = wbExcel.Worksheets("SALADES")
        wsExcel.Activate()
        NbSalades = wsExcel.Cells(1, 100).value - 1

        If NbSalades > 0 Then
            For j = 1 To NbSalades
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
            Next j

            Call ListeCategorie(NbSalades, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes laitages---------------------------------------
        wsExcel = wbExcel.Worksheets("LAITAGES")
        wsExcel.Activate()
        NbLaitages = wsExcel.Cells(1, 100).value - 1

        If NbLaitages > 0 Then
            For j = 1 To NbLaitages
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 5).value
            Next j

            Call ListeCategorie(NbLaitages, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes divers---------------------------------------
        wsExcel = wbExcel.Worksheets("DIVERS")
        wsExcel.Activate()
        wsExcel.Cells(1, 100).value = "=COUNTA(C[-99])"
        NbDivers = wsExcel.Cells(1, 100).value - 1

        If NbDivers > 0 Then
            For j = 1 To NbDivers
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 2).value
            Next j

            Call ListeCategorie(NbDivers, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        ' -----------------V�rification existence dans la base prix-------------------
        For i = 1 To NbCat
            Test = False
            For j = 1 To nbPrix
                If Categorie(i) = CodePrix(j) Then
                    Test = True
                    Exit For
                End If
            Next j
            If Test = False Then
                TexteMsg = "Cat�gorie " & Categorie(i) & " non d�clar�e dans la base PRIX"
                Call Reporting("PRIX", "ARRET", TexteMsg, "PRIX")
                Call Colexit()
                '  Console.WriteLine("ARRET DE L'APPLICATION, consulter le RAPPORT")
                Exit Sub
            End If
        Next i

        '-----------Formattage feuille AIDA-----------------------------------------------------
        If FeuilleExiste("AIDA") = True Then
            appExcel.Application.DisplayAlerts = False
            wbExcel.Worksheets("AIDA").delete()
            appExcel.DisplayAlerts = True

        End If
        wsExcel = wbExcel.Worksheets.Add()
        wsExcel.Name = "AIDA"


        '--------recopie colonnes familles ---------------------------------------------------
        wsExcel = wbExcel.Worksheets("RESULTATS")
        CType(wsExcel.Columns("A:C"), Range).Copy()

        wsExcel = wbExcel.Worksheets("AIDA")
        CType(wsExcel.Columns("A:C"), Range).PasteSpecial(Excel.XlPasteType.xlPasteAll)
        CType(wsExcel.Rows(1), Range).Insert(XlDirection.xlUp)
        CType(wsExcel.Rows(1), Range).RowHeight = 164

        '---------d�cale l'ent�te des 3 premieres colonnes -----------------------------
        wsExcel.Range("A2:C2").Copy()
        wsExcel.Range("A1:C1").PasteSpecial(Excel.XlPasteType.xlPasteAll)
        wsExcel.Range("A2:C2").Clear()
        CType(wsExcel.Columns("B:B"), Range).ColumnWidth = 16

        CType(wsExcel.Rows(2), Range).RowHeight = 15

        AlphaColTri = AlphaCol(NbCat + 4)
        With wsExcel.Range("D:" & AlphaColTri)
            .ColumnWidth = 13
            .HorizontalAlignment = -4108
        End With
        wsExcel.Cells(2, NbCat + 4).value = "PRIX TOTAL"

        Decal = 3
        For j = 1 To NbCat
            wsExcel.Cells(2, Decal + j).value = Categorie(j) & " " & UnAIDA(j)
        Next

        '------------------V�rification des totaux---------------------------------
        'Test = False
        wsExcel = wbExcel.Worksheets("RESULTATS")
        Decal = 5
        If nbdenrees > 0 Then
            Call TestSomme(nbdenrees)
            Decal += nbdenrees + 3
        End If
        If NbPreparations > 0 Then
            Call TestSomme(NbPreparations)
            Decal += NbPreparations + 1
        End If
        If NbSalades > 0 Then
            Call TestSomme(NbSalades)
            Decal += NbSalades + 1
        End If
        If NbLaitages > 0 Then
            Call TestSomme(NbLaitages)
            Decal += NbLaitages + 2
        End If
        If NbDivers > 0 Then
            Call TestSomme(NbDivers)
        End If

        '------------Report des cumuls par cat�gorie, pour chaque famille---------------------
        For i = 1 To NbFamilles
            For k = 1 To NbCat
                PoidsCat(k) = 0
            Next k

            Decal = 5
            Erreur = False
            '----------------------------------------------------------
            If nbdenrees > 0 Then
                wsExcel = wbExcel.Worksheets("VIANDES")
                For j = 1 To nbdenrees
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
                    Poids(j) = wsExcel.Cells(j + 1, 2).value / 1000
                Next j
                Call ReportCumul(nbdenrees, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv, Erreur)
                Decal = Decal + nbdenrees + 3
                If Erreur Then
                    Call Colexit()
                    Exit Sub
                End If
            End If
            '--------------------------------------------------------------
            If NbPreparations > 0 Then
                wsExcel = wbExcel.Worksheets("PREPARATIONS")
                For j = 1 To NbPreparations
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
                    Poids(j) = wsExcel.Cells(j + 1, 7).value
                Next j
                Call ReportCumul(NbPreparations, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv, Erreur)
                Decal = Decal + NbPreparations + 1
                If Erreur Then
                    Call Colexit()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------
            If NbSalades > 0 Then
                wsExcel = wbExcel.Worksheets("SALADES")
                For j = 1 To NbSalades
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).value
                    Poids(j) = wsExcel.Cells(j + 1, 7).value
                Next j
                Call ReportCumul(NbSalades, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv, Erreur)
                Decal = Decal + NbSalades + 1
                If Erreur Then
                    Call Colexit()
                    Exit Sub
                End If
            End If
            '---------------------------------------------------------------------
            If NbLaitages > 0 Then
                wsExcel = wbExcel.Worksheets("LAITAGES")
                For j = 1 To NbLaitages
                    Equiv(j) = wsExcel.Cells(j + 1, 4).value
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 5).value
                    Poids(j) = wsExcel.Cells(j + 1, 6).value
                Next j
                Call ReportCumul(NbLaitages, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv, Erreur)
                Decal = Decal + NbLaitages + 2
                If Erreur Then
                    Call Colexit()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------
            If NbDivers > 0 Then
                wsExcel = wbExcel.Worksheets("DIVERS")
                For j = 1 To NbDivers
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 2).value
                    Poids(j) = wsExcel.Cells(j + 1, 3).value
                Next j
                Call ReportCumul(NbDivers, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv, Erreur)
                If Erreur Then
                    Call Colexit()
                    Exit Sub
                End If
            End If
            '----------------------------------------------------------------------------
            wsExcel = wbExcel.Worksheets("AIDA")        'Ecriture des r�sultats dans l'onglet AIDA
            For k = 1 To NbCat
                wsExcel.Cells(i + 2, k + 3).value = PoidsCat(k)
            Next k
        Next i

        '------------Ajustement des arrondis pour l'unit� KgM-------------------------------
        wsExcel = wbExcel.Worksheets("AIDA")
        For k = 1 To NbCat
            If UnAIDA(k) = "[kgM]" Then
                SousTotal = 0
                TotalArrondi = 0
                For i = 1 To NbFamilles
                    SousTotal += wsExcel.Cells(i + 2, k + 3).value               'calcule le total brut 
                    Arrondi = Math.Round(wsExcel.Cells(i + 2, k + 3).value, 2)   ' carrondi les valeurs
                    wsExcel.Cells(i + 2, k + 3).value = Arrondi
                    TotalArrondi += wsExcel.Cells(i + 2, k + 3).value           ' calcule le total des arrondis
                Next i
                For i = 1 To NbFamilles
                    If wsExcel.Cells(i + 2, k + 3).value > 0 Then
                        Ecart = SousTotal - TotalArrondi            ' calcule la diff�rence entre les deux sous-totaux
                        wsExcel.Cells(i + 2, k + 3).value += Ecart  ' reporte l'�cart entre le total brut et le total des arrondis
                        Exit For            ' sort de la boucle d�s que l'�cart est report�
                    End If
                Next i
            End If
        Next k

        '-----------------------Prix du panier -------------------------------------------
        For i = 1 To NbFamilles
            PrixPanier = 0
            For k = 1 To NbCat
                PrixPanier += wsExcel.Cells(i + 2, k + 3).value * PrixListe(k)
            Next k
            wsExcel.Cells(i + 2, NbCat + 4).value = PrixPanier
        Next i

        '------------------Mise en forme onglet AIDA------------------------------------------

        i = 2
        AlphaColTri = AlphaCol(NbCat + 3)
        While i < NbFamilles + 3         'colorie les lignes une sur deux
            wsExcel.Range("A" & i & ":" & AlphaColTri & i).Interior.ColorIndex = 19
            i += 2
        End While

        '---formatage des colonnes en kilo ou autre-------------------------------------------------
        For i = 1 To NbCat
            AlphaColTri = AlphaCol(i + 3)
            If UnAIDA(i) = "[kgC]" Or UnAIDA(i) = "[kgM]" Then
                FormatCol = "###0.0#;;#"
                CType(wsExcel.Columns(AlphaColTri & ":" & AlphaColTri), Range).NumberFormat = FormatCol
            Else
                FormatCol = "####;;#"
                CType(wsExcel.Columns(AlphaColTri & ":" & AlphaColTri), Range).NumberFormat = FormatCol
            End If
        Next

        AlphaColTri = AlphaCol(NbCat + 4)       'colonne Totaux
        FormatCol = "#0.0#;;#"
        CType(wsExcel.Columns(AlphaColTri & ":" & AlphaColTri), Range).NumberFormat = FormatCol

        '--------------Codes barres ------------------------------------------------
        Decal = 3
        For j = 1 To NbCat
            jdec = j + Decal
            Contenu = CatAIDA(j)

            Call CodeBarreBMP(j, Contenu)

            wsExcel = wbExcel.Worksheets("AIDA")
            wsExcel.Activate()
            wsExcel.Cells(1, jdec).select

            Dim dT, dL, dW, dH As Single
            dT = wsExcel.Cells(1, jdec).Top          'coordonn�es du haut de l'image
            dL = wsExcel.Cells(1, jdec).left         ' coordonn�es du cot� gauche de l'image
            dW = 70                                  ' largeur de l'image
            dH = 160                                 ' hauteur de l'image

            wsExcel.Shapes.AddPicture(CheminBureau + "\Image" & j & ".bmp",
                Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, dL, dT, dW, dH)
        Next

        For j = 1 To NbCat
            Microsoft.VisualBasic.FileSystem.Kill(CheminBureau + "\Image" & j & ".bmp")
        Next

        Call Colexit()

    End Sub

    Sub TestSomme(nbden As Integer)
        '  ----- test si une colonne est vide ------------------
        Dim j As Integer
        Dim AlphaColTri As String

        For j = 1 To nbden
            If wsExcel.Cells(NbFamilles + 2, j + Decal).value = 0 Then
                AlphaColTri = AlphaCol(j + Decal)
                TexteMsg = "Ligne " & NbFamilles + 2 & "  La somme de la colonne " & AlphaColTri & " est nulle"
                Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RAPPORT")
            End If
        Next j

    End Sub
    Sub ReportCumul(nbdenrees As Integer, Decal As Integer, i As Integer, CodePrixDenree() As String, Categorie() As String,
    PoidsCat() As Single, Poids() As Single, UnAIDA() As String, Equiv() As Single, Erreur As Boolean)

        ' relecture des resultats, calcul des cumuls, denr�e par denr�e, � chaque cat�gorie en fonction de l'unit�
        ' si l'unit� n'est pas reconnue, sortie sous Erreur = True

        Dim j As Integer
        Dim k As Integer
        Dim CatErreur As String = ""
        Dim UnitErreur As String = ""

        wsExcel = wbExcel.Worksheets("RESULTATS")
        For j = 1 To nbdenrees
            For k = 1 To NbCat
                Select Case UCase(UnAIDA(k))
                    Case "[BOI]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).value * Equiv(j)
                    Case "[KGC]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + Poids(j) * wsExcel.Cells(i + 1, j + Decal).value
                    Case "[KGM]"
                        If Categorie(k) = CodePrixDenree(j) And wsExcel.Cells(NbFamilles + 2, j + Decal).value <> 0 Then
                            PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).value * Poids(j) / wsExcel.Cells(NbFamilles + 2, j + Decal).value
                        End If
                    Case "[UN]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).value
                    Case Else
                        Erreur = True
                        CatErreur = Categorie(k)
                        UnitErreur = UnAIDA(k)
                End Select
            Next k
            If Erreur Then
                TexteMsg = "ReportCumul: Pour la cat�gorie " & CatErreur & ", l'unit� " & UnitErreur & " n'est pas reconnue"
                Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RAPPORT")
                Exit Sub
            End If

        Next j
    End Sub
    Sub ListeCategorie(nbdenrees As Integer, nbPrix As Integer, CodePrixDenree() As String, CodePrix() As String,
    CodeAIDA() As String, UnitAIDA() As String, Categorie() As String, CatAIDA() As String,
    UnAIDA() As String, PrixAIDA() As Single, PrixListe() As Single)

        '  Construction de la liste unique des cat�gories � partir des cat�gories utilis�es dans les diff�rents onglets de denr�es

        Dim j As Integer
        Dim k As Integer
        Dim Test As Boolean

        ' initialisation de la premi�re valeur
        If NbCat = 0 Then
            Categorie(1) = CodePrixDenree(1)
            For k = 1 To nbPrix
                If Categorie(1) = CodePrix(k) Then
                    CatAIDA(1) = CodeAIDA(k)
                    UnAIDA(1) = UnitAIDA(k)
                    PrixListe(1) = PrixAIDA(k)
                    Exit For
                End If
            Next k
            NbCat = 1
        End If

        'teste si la cat�gorie est d�ja dans la liste et, si non, ajoute la nouvelle cat�gorie � la liste
        For j = 1 To nbdenrees
            Test = False

            For k = 1 To NbCat
                If CodePrixDenree(j) = Categorie(k) Then Test = True
            Next k

            If Test = False Then
                NbCat += 1
                Categorie(NbCat) = CodePrixDenree(j)
                For k = 1 To nbPrix
                    If Categorie(NbCat) = CodePrix(k) Then
                        CatAIDA(NbCat) = CodeAIDA(k)
                        UnAIDA(NbCat) = UnitAIDA(k)
                        PrixListe(NbCat) = PrixAIDA(k)
                        Exit For
                    End If
                Next k
            End If

        Next j

    End Sub

    Public Sub CodeBarreBMP(j As Integer, Contenu As String)
        '*******************************************************************************************************************
        'Variables envoy�es lors de l'appel de la routine
        '   j = index de boucle pour sauvegarder l'image, permet d'enchainer plusieurs images � la suite
        '   Contenu   = String du code barre � encoder
        '******************************************************************************************************************
        ' G�n�ration d'un code barre 128 dans un fichier image bitmap
        '       --------------------------------------------------------------------
        '       Bas� sur le code de  Dominique KIRCHHOFER  "Access : cr�er des codes-barres 128 en VBA"  sur Developpez.com
        '       pour les routines  Code128 et MotifCodeBarre128
        '       --------------------------------------------------------------------
        '       Adapt� � VBA Excel en 2011:
        '           g�n�re une zone de texte dans la feuille de calcul
        '           appel direct � la routine CodeBarre � ins�rer dans un module
        '           Pas de twips en Excel, travail en Point
        '       ---------------------------------------------------------------------
        '       Traduction en VB.NET Juin 2024,
        '           cr�ation d'une image bitmap � partir de GDI+  (plus rapide car �vite les Shapes pour tracer les lignes)
        '           sauvegarde de l'image bitmap sur la directory de l'application
        '******************************************************************************************************************

        'Variables locales
        Dim strChaine As String                 ' Variable recevant le code 128, apr�s encodage de la cha�ne de caract�res
        Dim strCaractere As String              ' Variable recevant successivement chaque caract�re du code 128, avant leur conversion
        Dim strBarres                           ' Variable recevant successivement les caract�res du code 128, apr�s conversion
        Dim strCodeBarres As String             ' Variable contenant le code 128 converti
        Dim i As Long                           ' Variable de compteur

        Dim lngLargeurCodeBarres As Integer      ' Largeur du code-barres
        Dim strTypeModule As String              ' Type d'un module : 1 = barre / 0 = espace
        Dim IntHauteur As Integer = 400          ' Hauteur de l'image
        Dim IntLargeur As Integer = 850          ' Largeur de l'image
        Dim IntHautModule As Integer = 300       ' hauteur des barres
        Dim IntLargModule As Integer = 5         ' largeur des barres
        Dim X1 As Integer
        Dim X2 As Integer
        Dim Y1 As Integer
        Dim Y2 As Integer
        Dim IntMargeGauche As Integer           ' marge entre le bord de l'image et le d�but du trac�

        strCodeBarres = ""

        '---------Appel routine d'encodage de la cha�ne de caract�res en code 128------------------------
        strChaine = Code128(Contenu)

        '---------Conversion des caract�res. Le chiffre "1" repr�sente les barres, le chiffre "0" les espaces
        For i = 1 To Len(strChaine)
            strCaractere = Mid(strChaine, i, 1)
            strBarres = MotifCodeBarres128(strCaractere)                        'Appel � la routine de conversion
            strCodeBarres &= strBarres
        Next i

        strCodeBarres = "00000000000" & strCodeBarres & "00000000000"           'Ajout des "Quiet zone" en d�but et en fin du code-barres
        lngLargeurCodeBarres = Len(strCodeBarres) * IntLargModule              'Largeur du code-barres, "Quiet zone" comprises

        ' -------redimensionne la taille de l'image si n�cessaire ---------------------------
        If IntLargeur < lngLargeurCodeBarres Then IntLargeur = CInt(lngLargeurCodeBarres * 1.1)
        IntMargeGauche = CInt((IntLargeur - lngLargeurCodeBarres) / 2)

        ' *******************************************************************
        ' cr�ation du graphique
        '********************************************************************
        Dim newBitmap As New Bitmap(IntLargeur, IntHauteur) 'cr�ons un BitMap
        Dim g As Graphics = Graphics.FromImage(newBitmap) 'cr�ons un Graphics et y mettre le BitMap

        Dim blackPen As New Pen(Color.Black, 5)    'cr�er un stylet noir d'�paisseur 
        ' Dim YeloPen As New Pen(Color.Yellow, 5)
        Dim WhitePen As New Pen(Color.White, 5)

        g.FillRectangle(New SolidBrush(Color.White), 0, 0, IntLargeur, IntHauteur)

        X1 = IntMargeGauche
        X2 = X1
        Y1 = 380
        Y2 = Y1 - IntHautModule

        '------Cr�ation de la zone de texte ------------------------------------------

        Dim drawFont As New System.Drawing.Font("Arial", 35)
        Dim drawBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Black)

        g.DrawString(Contenu, drawFont, drawBrush, 250, 23)

        '-----Tra�age du code-barres----------------------------------------------
        For i = 1 To Len(strCodeBarres)
            strTypeModule = Mid(strCodeBarres, i, 1)                                'Type de module, barre ou espace, � tracer
            Select Case strTypeModule
                Case "1"
                    g.DrawLine(blackPen, X1, Y1, X2, Y2)
                    X1 += IntLargModule
                    X2 = X1

                Case "0"
                    g.DrawLine(WhitePen, X1, Y1, X2, Y2)
                    X1 += IntLargModule
                    X2 = X1
            End Select

        Next i

        'PictureBox1.Image = newBitmap
        newBitmap.RotateFlip(RotateFlipType.Rotate270FlipNone)      'Rotation de l'image verticalement si besoin
        'PictureBox1.Image.Save(Cheminbureau + "\Image" & j & ".bmp")
        newBitmap.Save(CheminBureau + "\Image" & j & ".bmp")

    End Sub
    Public Function Code128(strChaine As String) As String

        'Caract�re en cours de traitement
        Dim strCaractere As String
        'Cha�ne de caract�res temporaire
        Dim strChaineTemp As String
        'Caract�re temporaire en cours de traitement
        Dim strCarTemp As String
        'Table utilis�e (table B)
        Dim blnTableB As Boolean
        'Table utilis�e (table C)
        Dim blnTableC As Boolean
        'Valeur de la cl� de contr�le
        Dim lngCaractereControle As Long
        'Variable de compteur
        Dim i As Long
        'Variable de compteur
        Dim j As Long

        'G�n�ration d'une erreur d�finie par l'utilisateur
        If String.IsNullOrEmpty(strChaine) Then
            Call Reporting("AIDA", "ARRET", "Cha�ne de caract�res Code-Barres nulle ", "AIDA")
            Code128 = ""
            ' MsgBox("chaine nulle ")
            Exit Function
        End If

        '--------initialisation-----------------------------------
        strChaineTemp = ""
        Code128 = ""

        'Codage de la cha�ne de caract�res
        For i = 1 To Len(strChaine)

            'Extraction d'un caract�re de la cha�ne
            strCaractere = Mid(strChaine, i, 1)

            'Ajout du caract�re � la cha�ne temporaire
            strChaineTemp &= strCaractere

            'D�but sur table B ou C
            If Not blnTableB And Not blnTableC Then

                'Quatre caract�res num�riques sont n�cessaires pour d�buter en table C
                If IsNumeric(strCaractere) Then

                    'La cha�ne temporaire contient quatre caract�res num�riques, d�but sur table C
                    If Len(strChaineTemp) = 4 Then

                        'Ajout du caract�re de d�marrage de la table C
                        Code128 = ChrW(210)

                        'Traitement de quatre caract�res. Ajout de deux caract�res optimis�s
                        For j = 1 To 4 Step 2
                            strCarTemp = Val(Mid(strChaineTemp, j, 2))
                            strCarTemp = IIf(strCarTemp < 95, strCarTemp + 32, strCarTemp + 105)
                            Code128 &= ChrW(strCarTemp)
                        Next j

                        'Purge de la cha�ne de caract�res temporaire
                        strChaineTemp = ""

                        'La table C est utilis�e
                        blnTableC = True

                    End If

                    'Le nombre de caract�res num�riques en d�but de cha�ne est inf�rieur � quatre, d�but sur table B
                Else

                    'Ajout du caract�re de d�marrage de la table B
                    Code128 &= ChrW(209)

                    'Ajout des caract�res de la cha�ne temporaire
                    For j = 1 To Len(strChaineTemp)
                        Code128 &= Mid(strChaineTemp, j, 1)
                    Next j

                    'Purge de la cha�ne de caract�res temporaire
                    strChaineTemp = ""

                    'La table B est utilis�e
                    blnTableB = True

                End If

                'Traitement de la suite de la cha�ne de caract�res
            Else

                'Traitement sur table C, tentative de traiter des caract�res num�riques suppl�mentaires
                If blnTableC Then

                    'Deux caract�res num�riques sont n�cessaires pour continuer sur table C
                    If IsNumeric(strCaractere) Then

                        'La chaine temporaire contient deux caract�res num�riques
                        If Len(strChaineTemp) = 2 Then

                            'Traitement de deux caract�res. Ajout d'un caract�re optimis�
                            strCarTemp = Val(Mid(strChaineTemp, 1, 2))
                            strCarTemp = IIf(strCarTemp < 95, strCarTemp + 32, strCarTemp + 105)
                            Code128 &= ChrW(strCarTemp)

                            'Purge de la cha�ne de caract�res temporaire
                            strChaineTemp = ""

                        End If

                        'Le nombre de caract�res num�riques est inf�rieur � deux
                    Else

                        'Permutation sur table B
                        Code128 &= ChrW(205)

                        'Ajout des caract�res de la cha�ne temporaire
                        For j = 1 To Len(strChaineTemp)
                            Code128 &= Mid(strChaineTemp, j, 1)
                        Next j

                        'Purge de la cha�ne de caract�res temporaire
                        strChaineTemp = ""

                        'La table B est utilis�e
                        blnTableC = False
                        blnTableB = True

                    End If

                    'Traitement sur table B, tentative de permuter sur table C pour optimiser le code
                Else

                    'Le caract�re est num�rique
                    If IsNumeric(strCaractere) Then

                        'Si le reste de la cha�ne et le contenu de la cha�ne temporaire est �gal
                        '� au moins six caract�res
                        If Len(strChaine) - i + Len(strChaineTemp) >= 6 Then

                            'La cha�ne temporaire contient six caract�res num�riques
                            If Len(strChaineTemp) = 6 Then

                                'Permutation sur table C
                                Code128 &= ChrW(204)

                                'Traitement de six caract�res num�riques. Ajout de trois caract�res optimis�s
                                For j = 1 To 6 Step 2
                                    strCarTemp = Val(Mid(strChaineTemp, j, 2))
                                    strCarTemp = IIf(strCarTemp < 95, strCarTemp + 32, strCarTemp + 105)
                                    Code128 &= ChrW(strCarTemp)
                                Next j

                                'Purge de la cha�ne de caract�res temporaire
                                strChaineTemp = ""

                                'La table C est utilis�e
                                blnTableB = False
                                blnTableC = True

                            End If

                            'Le nombre de caract�res de la cha�ne temporaire et le reste de caract�res � traiter est inf�rieur � six
                        Else

                            'Le reste de caract�res � traiter est �gal � cinq
                            If Len(strChaine) - i + 1 = 5 Then

                                'Ajout du caract�re sur table B
                                Code128 &= strChaineTemp

                                'Purge de la cha�ne de caract�res temporaire
                                strChaineTemp = ""

                            End If

                            'Si le nombre de caract�res restant est �gal ou inf�rieur � quatre
                            If Len(strChaine) - i + 1 <= 4 Then

                                'La cha�ne temporaire contient quatre caract�res num�riques
                                If Len(strChaineTemp) = 4 Then

                                    'Permutation sur table C
                                    Code128 &= ChrW(204)

                                    'Traitement de quatre caract�res num�riques. Ajout de deux caract�res optimis�s
                                    For j = 1 To 4 Step 2
                                        strCarTemp = Val(Mid(strChaineTemp, j, 2))
                                        strCarTemp = IIf(strCarTemp < 95, strCarTemp + 32, strCarTemp + 105)
                                        Code128 &= ChrW(strCarTemp)
                                    Next j

                                    'Purge de la cha�ne de caract�res temporaire
                                    strChaineTemp = ""

                                End If

                            End If

                        End If

                        'Le caract�re en cours n'est pas num�rique
                    Else

                        'Ajout du caract�re sur table B
                        For j = 1 To Len(strChaineTemp)
                            Code128 &= Mid(strChaineTemp, j, 1)
                        Next j

                        'Purge de la cha�ne de caract�res temporaire
                        strChaineTemp = ""

                    End If

                End If

            End If

            'Traitement du dernier caract�re de la cha�ne
            If i = Len(strChaine) And Len(strChaineTemp) >= 1 Then

                'La table C est en cours d'utilisation
                If blnTableC Then

                    'Permutation vers la table B
                    Code128 &= ChrW(205)

                    'Ajout du dernier caract�re sur table B
                    Code128 &= strChaineTemp

                    'La table B est en cours d'utilisation
                ElseIf blnTableB Then

                    'Ajout des caract�res de la cha�ne temporaire
                    For j = 1 To Len(strChaineTemp)
                        Code128 &= Mid(strChaineTemp, j, 1)
                    Next j

                    'Aucune des deux tables n'est utilis�e. La cha�ne de caract�res contient moins
                    'de quatre caract�res num�riques
                Else

                    'D�but sur table B
                    Code128 &= ChrW(209)

                    'Ajout des caract�res de la cha�ne temporaire
                    For j = 1 To Len(strChaineTemp)
                        Code128 &= Mid(strChaineTemp, j, 1)
                    Next j

                End If

            End If

        Next i

        'Calcul de la valeur de la cl� de contr�le
        For j = 1 To Len(Code128)
            strCarTemp = AscW(Mid(Code128, j, 1))
            strCarTemp = IIf(strCarTemp < 127, strCarTemp - 32, strCarTemp - 105)
            If j = 1 Then lngCaractereControle = strCarTemp
            lngCaractereControle = (lngCaractereControle + (j - 1) * strCarTemp) Mod 103
        Next

        'Caract�re ASCII de la cl� de contr�le
        lngCaractereControle = IIf(lngCaractereControle < 95, lngCaractereControle + 32, lngCaractereControle + 105)

        'Ajout du caract�re ASCII de la cl� de contr�le et du caract�re d'arr�t
        Code128 = Code128 & ChrW(lngCaractereControle) & ChrW(211)

        Exit Function

    End Function

    Public Function MotifCodeBarres128(strChaine As String) As String

        '  On Error GoTo GestionErreurs

        Select Case AscW(strChaine)
            Case 32 : MotifCodeBarres128 = "11011001100" ' Caract�re = Espace / Table B = Espace / Table C = 00
            Case 33 : MotifCodeBarres128 = "11001101100" ' Caract�re = ! / Table B = ! / Table C = 01
            Case 34 : MotifCodeBarres128 = "11001100110" ' Caract�re = " / Table B = " / Table C = 02
            Case 35 : MotifCodeBarres128 = "10010011000" ' Caract�re = # / Table B = # / Table C = 03
            Case 36 : MotifCodeBarres128 = "10010001100" ' Caract�re = $ / Table B = $ / Table C = 04
            Case 37 : MotifCodeBarres128 = "10001001100" ' Caract�re = % / Table B = % / Table C = 05
            Case 38 : MotifCodeBarres128 = "10011001000" ' Caract�re = & / Table B = & / Table C = 06
            Case 39 : MotifCodeBarres128 = "10011000100" ' Caract�re = ' / Table B = ' / Table C = 07
            Case 40 : MotifCodeBarres128 = "10001100100" ' Caract�re = ( / Table B = ( / Table C = 08
            Case 41 : MotifCodeBarres128 = "11001001000" ' Caract�re = ) / Table B = ) / Table C = 09
            Case 42 : MotifCodeBarres128 = "11001000100" ' Caract�re = * / Table B = * / Table C = 10
            Case 43 : MotifCodeBarres128 = "11000100100" ' Caract�re = + / Table B = + / Table C = 11
            Case 44 : MotifCodeBarres128 = "10110011100" ' Caract�re = , / Table B = , / Table C = 12
            Case 45 : MotifCodeBarres128 = "10011011100" ' Caract�re = - / Table B = - / Table C = 13
            Case 46 : MotifCodeBarres128 = "10011001110" ' Caract�re = . / Table B = . / Table C = 14
            Case 47 : MotifCodeBarres128 = "10111001100" ' Caract�re = / / Table B = / / Table C = 15
            Case 48 : MotifCodeBarres128 = "10011101100" ' Caract�re = 0 / Table B = 0 / Table C = 16
            Case 49 : MotifCodeBarres128 = "10011100110" ' Caract�re = 1 / Table B = 1 / Table C = 17
            Case 50 : MotifCodeBarres128 = "11001110010" ' Caract�re = 2 / Table B = 2 / Table C = 18
            Case 51 : MotifCodeBarres128 = "11001011100" ' Caract�re = 3 / Table B = 3 / Table C = 19
            Case 52 : MotifCodeBarres128 = "11001001110" ' Caract�re = 4 / Table B = 4 / Table C = 20
            Case 53 : MotifCodeBarres128 = "11011100100" ' Caract�re = 5 / Table B = 5 / Table C = 21
            Case 54 : MotifCodeBarres128 = "11001110100" ' Caract�re = 6 / Table B = 6 / Table C = 22
            Case 55 : MotifCodeBarres128 = "11101101110" ' Caract�re = 7 / Table B = 7 / Table C = 23
            Case 56 : MotifCodeBarres128 = "11101001100" ' Caract�re = 8 / Table B = 8 / Table C = 24
            Case 57 : MotifCodeBarres128 = "11100101100" ' Caract�re = 9 / Table B = 9 / Table C = 25
            Case 58 : MotifCodeBarres128 = "11100100110" ' Caract�re = : / Table B = : / Table C = 26
            Case 59 : MotifCodeBarres128 = "11101100100" ' Caract�re = ; / Table B = ; / Table C = 27
            Case 60 : MotifCodeBarres128 = "11100110100" ' Caract�re = < / Table B = < / Table C = 28
            Case 61 : MotifCodeBarres128 = "11100110010" ' Caract�re = = / Table B = = / Table C = 29
            Case 62 : MotifCodeBarres128 = "11011011000" ' Caract�re = > / Table B = > / Table C = 30
            Case 63 : MotifCodeBarres128 = "11011000110" ' Caract�re = ? / Table B = ? / Table C = 31
            Case 64 : MotifCodeBarres128 = "11000110110" ' Caract�re = @ / Table B = @ / Table C = 32
            Case 65 : MotifCodeBarres128 = "10100011000" ' Caract�re = A / Table B = A / Table C = 33
            Case 66 : MotifCodeBarres128 = "10001011000" ' Caract�re = B / Table B = B / Table C = 34
            Case 67 : MotifCodeBarres128 = "10001000110" ' Caract�re = C / Table B = C / Table C = 35
            Case 68 : MotifCodeBarres128 = "10110001000" ' Caract�re = D / Table B = D / Table C = 36
            Case 69 : MotifCodeBarres128 = "10001101000" ' Caract�re = E / Table B = E / Table C = 37
            Case 70 : MotifCodeBarres128 = "10001100010" ' Caract�re = F / Table B = F / Table C = 38
            Case 71 : MotifCodeBarres128 = "11010001000" ' Caract�re = G / Table B = G / Table C = 39
            Case 72 : MotifCodeBarres128 = "11000101000" ' Caract�re = H / Table B = H / Table C = 40
            Case 73 : MotifCodeBarres128 = "11000100010" ' Caract�re = I / Table B = I / Table C = 41
            Case 74 : MotifCodeBarres128 = "10110111000" ' Caract�re = J / Table B = J / Table C = 42
            Case 75 : MotifCodeBarres128 = "10110001110" ' Caract�re = K / Table B = K / Table C = 43
            Case 76 : MotifCodeBarres128 = "10001101110" ' Caract�re = L / Table B = L / Table C = 44
            Case 77 : MotifCodeBarres128 = "10111011000" ' Caract�re = M / Table B = M / Table C = 45
            Case 78 : MotifCodeBarres128 = "10111000110" ' Caract�re = N / Table B = N / Table C = 46
            Case 79 : MotifCodeBarres128 = "10001110110" ' Caract�re = O / Table B = O / Table C = 47
            Case 80 : MotifCodeBarres128 = "11101110110" ' Caract�re = P / Table B = P / Table C = 48
            Case 81 : MotifCodeBarres128 = "11010001110" ' Caract�re = Q / Table B = Q / Table C = 49
            Case 82 : MotifCodeBarres128 = "11000101110" ' Caract�re = R / Table B = R / Table C = 50
            Case 83 : MotifCodeBarres128 = "11011101000" ' Caract�re = S / Table B = S / Table C = 51
            Case 84 : MotifCodeBarres128 = "11011100010" ' Caract�re = T / Table B = T / Table C = 52
            Case 85 : MotifCodeBarres128 = "11011101110" ' Caract�re = U / Table B = U / Table C = 53
            Case 86 : MotifCodeBarres128 = "11101011000" ' Caract�re = V / Table B = V / Table C = 54
            Case 87 : MotifCodeBarres128 = "11101000110" ' Caract�re = W / Table B = W / Table C = 55
            Case 88 : MotifCodeBarres128 = "11100010110" ' Caract�re = X / Table B = X / Table C = 56
            Case 89 : MotifCodeBarres128 = "11101101000" ' Caract�re = Y / Table B = Y / Table C = 57
            Case 90 : MotifCodeBarres128 = "11101100010" ' Caract�re = Z / Table B = Z / Table C = 58
            Case 91 : MotifCodeBarres128 = "11100011010" ' Caract�re = [ / Table B = [ / Table C = 59
            Case 92 : MotifCodeBarres128 = "11101111010" ' Caract�re = \ / Table B = \ / Table C = 60
            Case 93 : MotifCodeBarres128 = "11001000010" ' Caract�re = ] / Table B = ] / Table C = 61
            Case 94 : MotifCodeBarres128 = "11110001010" ' Caract�re = ^ / Table B = ^ / Table C = 62
            Case 95 : MotifCodeBarres128 = "10100110000" ' Caract�re = _ / Table B = _ / Table C = 63
            Case 96 : MotifCodeBarres128 = "10100001100" ' Caract�re = ` / Table B = ` / Table C = 64
            Case 97 : MotifCodeBarres128 = "10010110000" ' Caract�re = a / Table B = a / Table C = 65
            Case 98 : MotifCodeBarres128 = "10010000110" ' Caract�re = b / Table B = b / Table C = 66
            Case 99 : MotifCodeBarres128 = "10000101100" ' Caract�re = c / Table B = c / Table C = 67
            Case 100 : MotifCodeBarres128 = "10000100110" ' Caract�re = d / Table B = d / Table C = 68
            Case 101 : MotifCodeBarres128 = "10110010000" ' Caract�re = e / Table B = e / Table C = 69
            Case 102 : MotifCodeBarres128 = "10110000100" ' Caract�re = f / Table B = f / Table C = 70
            Case 103 : MotifCodeBarres128 = "10011010000" ' Caract�re = g / Table B = g / Table C = 71
            Case 104 : MotifCodeBarres128 = "10011000010" ' Caract�re = h / Table B = h / Table C = 72
            Case 105 : MotifCodeBarres128 = "10000110100" ' Caract�re = i / Table B = i / Table C = 73
            Case 106 : MotifCodeBarres128 = "10000110010" ' Caract�re = j / Table B = j / Table C = 74
            Case 107 : MotifCodeBarres128 = "11000010010" ' Caract�re = k / Table B = k / Table C = 75
            Case 108 : MotifCodeBarres128 = "11001010000" ' Caract�re = l / Table B = l / Table C = 76
            Case 109 : MotifCodeBarres128 = "11110111010" ' Caract�re = m / Table B = m / Table C = 77
            Case 110 : MotifCodeBarres128 = "11000010100" ' Caract�re = n / Table B = n / Table C = 78
            Case 111 : MotifCodeBarres128 = "10001111010" ' Caract�re = o / Table B = o / Table C = 79
            Case 112 : MotifCodeBarres128 = "10100111100" ' Caract�re = p / Table B = p / Table C = 80
            Case 113 : MotifCodeBarres128 = "10010111100" ' Caract�re = q / Table B = q / Table C = 81
            Case 114 : MotifCodeBarres128 = "10010011110" ' Caract�re = r / Table B = r / Table C = 82
            Case 115 : MotifCodeBarres128 = "10111100100" ' Caract�re = s / Table B = s / Table C = 83
            Case 116 : MotifCodeBarres128 = "10011110100" ' Caract�re = t / Table B = t / Table C = 84
            Case 117 : MotifCodeBarres128 = "10011110010" ' Caract�re = u / Table B = u / Table C = 85
            Case 118 : MotifCodeBarres128 = "11110100100" ' Caract�re = v / Table B = v / Table C = 86
            Case 119 : MotifCodeBarres128 = "11110010100" ' Caract�re = w / Table B = w / Table C = 87
            Case 120 : MotifCodeBarres128 = "11110010010" ' Caract�re = x / Table B = x / Table C = 88
            Case 121 : MotifCodeBarres128 = "11011011110" ' Caract�re = y / Table B = y / Table C = 89
            Case 122 : MotifCodeBarres128 = "11011110110" ' Caract�re = z / Table B = z / Table C = 90
            Case 123 : MotifCodeBarres128 = "11110110110" ' Caract�re = { / Table B = { / Table C = 91
            Case 124 : MotifCodeBarres128 = "10101111000" ' Caract�re = | / Table B = | / Table C = 92
            Case 125 : MotifCodeBarres128 = "10100011110" ' Caract�re = } / Table B = } / Table C = 93
            Case 126 : MotifCodeBarres128 = "10001011110" ' Caract�re = ~ / Table B = ~ / Table C = 94
            Case 200 : MotifCodeBarres128 = "10111101000" ' Caract�re = � - Table B = Del / Table C = 95
            Case 201 : MotifCodeBarres128 = "10111100010" ' Caract�re = � / Table B = Fnc 3 / Table C = 96
            Case 202 : MotifCodeBarres128 = "11110101000" ' Caract�re = � / Table B = Fnc 2 / Table C = 97
            Case 203 : MotifCodeBarres128 = "11110100010" ' Caract�re = � / Table B = Shift / Table C = 98
            Case 204 : MotifCodeBarres128 = "10111011110" ' Caract�re = � - Table B = Table C / Table C = 99
            Case 205 : MotifCodeBarres128 = "10111101110" ' Caract�re = � - Table B = Fnc 4 / Table C = Table B
            Case 206 : MotifCodeBarres128 = "11101011110" ' Caract�re = � - Table B = Table A / Table C = Table A
            Case 207 : MotifCodeBarres128 = "11110101110" ' Caract�re = � - Table B = Fnc 1 / Table C = Fnc 1
            Case 208 : MotifCodeBarres128 = "11010000100" ' Caract�re = � - Table B = Start A / Table C = Start A
            Case 209 : MotifCodeBarres128 = "11010010000" ' Caract�re = � - Table B = Start B / Table C = Start B
            Case 210 : MotifCodeBarres128 = "11010011100" ' Caract�re = � - Table B = Start C / Table C = Start C
            Case 211 : MotifCodeBarres128 = "1100011101011" ' Caract�re = � - Table B = Stop / Table C = Stop
                'Erreur
            Case Else
                'MsgBox("motif inconnu " & AscW(strChaine))
                Call Reporting("AIDA", "ARRET", "Motif Code-Barre inconnu: " & AscW(strChaine), "AIDA")

        End Select

        Exit Function

        'GestionErreurs:

        'Transmet l'erreur � la proc�dure appelante
        'Err.Raise(Err.Number, "MotifCodeBarres128")

    End Function
End Module

