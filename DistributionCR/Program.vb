
Option Explicit On
Imports System.Globalization
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.ExcelErrorValue
Imports OfficeOpenXml.Sorting
Imports OfficeOpenXml.Style
Imports SkiaSharp

Public Module Program
    'Arguments d'entr�e
    Private argFileName As String
    Private argMode As Integer

    'D�claration des variables  Excel
    Dim appExcel As ExcelPackage          'Application Excel
    Dim wbExcel As ExcelWorkbook          'Classeur Excel
    Dim wsExcel As ExcelWorksheet         'Feuille Excel
    Dim StrInput As String                'nom complet du fichier de donn�es

    ' dimensions des variables (� changer en cas de besoin)
    Public Const MaxFamilles As Integer = 150
    Public Const MaxPrix As Integer = 300
    Public Const MaxDenrees As Integer = 100
    ' constantes de poids
    Public Const CoefPrepa As Integer = 4
    Public Const CoefSalade As Integer = 4

    ' variables de travail
    Dim Decal As Integer                        'd�calage de colonnes dans RESULTATS, 
    Dim nbReport As Integer                     'compteur de lignes dans RAPPORT
    Dim TexteMsg As String                      'variable de texte
    Dim TestErreur As Boolean                   ' Test si sortie sur erreur, arr�t du programme
    Dim NbErreur As Integer                     ' Compteur d'erreurs en Alerte
    Dim NbCat As Integer                         'nombre de cat�gories AIDA
    Dim CheminBureau As String                   'chemin pour sauvegarde de l'image code-barre
    Dim ValeurR As Single                        ' Valeur R�elle de retour de la sub VerifNumeric
    Dim ValeurI As Integer                       ' Valeur enti�re de retour de la sub  VerifNumeric

    'D�claration variables FAMILLES
    Dim NumFamille(MaxFamilles) As Integer       'Numero de caisse
    Dim NomFamille(MaxFamilles) As String        'Nom de famille
    Dim NBenef(MaxFamilles) As Integer           'Nombre de b�n�ficiaires
    Dim SansCochon(MaxFamilles) As Boolean       'indicateur Sans cochon
    Dim SansViande(MaxFamilles) As Boolean       'indicateur vegan
    Dim Panier(MaxFamilles) As Single            'panier de la famille
    Dim PoidsTheo(MaxFamilles) As Single         'poids th�orique de viande
    Dim PanierZeu(MaxFamilles) As Single         'panier d'oeufs
    Dim PoidsTheozeu(MaxFamilles) As Single      'quantit� theo d'oeufs
    Dim TestSCSV(MaxFamilles) As Integer         'cocatenation des indicateurs sanscochon et sansviande
    Dim NbFamilles As Integer                    'nombre de familles

    Sub Main()
        Dim NbLignes As Integer

        'Initialisation de EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Call ReadArgs()

        'Ouverture de l'application Excel
        appExcel = (New ExcelPackage(New FileInfo(argFileName)))
        wbExcel = appExcel.Workbook

        If FeuilleExiste("RAPPORT") = True Then
            wsExcel = wbExcel.Worksheets("RAPPORT")
            NbLignes = GetNonEmptyRows()   'Compte le nbre de lignes
            For i = 2 To NbLignes
                wsExcel.Cells(i, 1).Value = Nothing
                wsExcel.Cells(i, 2).Value = Nothing
                wsExcel.Cells(i, 8).Value = Nothing
            Next
        Else
            wsExcel = wbExcel.Worksheets.Add("RAPPORT")
            wsExcel.Name = "RAPPORT"

            wsExcel.Cells(1, 1).Value = "FEUILLE"
            wsExcel.Cells(1, 2).Value = "Probl�me"
            wsExcel.Cells(1, 8).Value = "Criticit�"
            For i = 1 To 7
                wsExcel.Columns(i).Width = 15
            Next
            With wsExcel.Cells("A1:H1").Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 43
                .Font.Bold = True
                .Border.Bottom.Style = ExcelBorderStyle.Medium
            End With
        End If

        nbReport = 1

        Select Case argMode
            Case 1
                Call Repartition()
                If TestErreur = False Then Call AIDA()      'Si des erreurs sont d�tect�es, on n'appelle pas AIDA
                Call Colexit()
            Case 2
                Call MAJ()
                If TestErreur = False Then Call AIDA()      'Si des erreurs sont d�tect�es, on n'appelle pas AIDA
                Call Colexit()
            Case Else
                Call Colexit()
        End Select
    End Sub

    Public Sub ReadArgs()
        Dim cliArgs = Environment.GetCommandLineArgs()
        If cliArgs.Length = 3 Then
            argFileName = cliArgs.GetValue(1)
            argMode = CInt(cliArgs.GetValue(2))
            Return
        End If

        'demande � l'utilisateur
        Console.WriteLine("  D I S T R I B U T I O N    C R O I X - R O U G E")
        Console.WriteLine("******************************************************")
        Console.WriteLine("Donner le chemin reseau du fichier")
        argFileName = Console.ReadLine()

        Console.WriteLine()
        Console.WriteLine("Choisir l'option de calcul:")
        Console.WriteLine("R�partition + A�da: tapez 1")
        Console.WriteLine("Mise � jour + A�da: tapez 2")
        Console.WriteLine()

        Dim StrOption As String = Console.ReadLine()
        Dim TestOption As Boolean = False

        Do Until TestOption = True
            Select Case StrOption
                Case "1"
                    argMode = 1
                    TestOption = True
                Case "2"
                    argMode = 2
                    TestOption = True
                Case "Exit"
                    argMode = 0
                    TestOption = True
                Case Else
                    Console.WriteLine("Option non reconnue, tapez 1 ou 2")
                    Console.WriteLine("Pour arr�ter, tapez Exit")
                    StrOption = Console.ReadLine()
            End Select
        Loop
    End Sub

    Public Sub Colexit()
        ' ------- Sauvegarde et fermeture d'Excel -------------------------------------
        If Not IsNothing(appExcel) Then
            appExcel.Save()
            appExcel = Nothing
            wbExcel = Nothing
        End If
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
        Dim col1 As Integer
        Dim col2 As Integer
        Dim col3 As Integer
        Dim Mode1 As Integer
        Dim Mode2 As Integer
        Dim Mode3 As Integer
        Dim A1 As New Random
        Dim ParamEcart As Integer
        Dim Saut As Boolean
        Dim Ecart As Single
        Dim EcartMaxi As Single
        Dim NumeMaxi As Integer
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
        '--------D�claration variables Anticipes
        Dim NomAnti(MaxDenrees) As String
        Dim TypeAnti(MaxDenrees) As String
        Dim DescAnti(MaxDenrees) As String
        Dim TailleAnti(MaxDenrees) As String
        Dim PoidsAnti(MaxDenrees) As Single
        Dim QuantAnti(MaxDenrees) As Integer
        Dim CodePrixanti(MaxDenrees) As String
        Dim Nbanti As Integer
        Dim TestNom As String
        Dim TestType As String
        Dim NbDenreesAnti As Integer
        Dim NbPrepaAnti As Integer
        Dim NbSaladeAnti As Integer
        Dim DescPrecedent As String
        Dim PoidsPrecedent As Single
        Dim QTotalAnti As Integer
        Dim SousTotal As Integer
        '---------D�claration variables PRIX------------------------
        Dim LibellePrix(MaxPrix) As String
        Dim CodePrix(MaxPrix) As String
        Dim CodeAida(MaxPrix) As String
        Dim NbPrix As Integer

        '---------Onglet DIVERS--------------------------------------
        Dim NbDivers As Integer
        Dim Divers(MaxDenrees) As String
        Dim TotalFamille As Single

        'Dim StrOption As String
        '**************************************************************************************
        '   Lecture des prix
        '**************************************************************************************
        NbErreur = 0
        TestErreur = False                  'initialisation du TestErreur � false

        If FeuilleExiste("PRIX") = False Then           'TEST pr�sence feuille
            Call Reporting("PRIX", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        wsExcel = wbExcel.Worksheets("PRIX")

        NbPrix = GetNonEmptyRows() - 1  'Compte le nbre de lignes

        If NbPrix > MaxPrix Then            'Test d�passement dimension maxi
            Call Reporting("PRIX", "ARRET", "Nombre de prix d�passe la dimension > " & MaxPrix, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        If NbPrix > 0 Then                  'D�but de lecture des donn�es 
            For i = 1 To NbPrix
                LibellePrix(i) = wsExcel.Cells(i + 1, 1).Value
                CodePrix(i) = wsExcel.Cells(i + 1, 2).Value
                CodeAida(i) = wsExcel.Cells(i + 1, 3).Value
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
            NbErreur += 1
        End If

        '*****************************************************************
        '   Lecture des viandes
        '******************************************************************

        If FeuilleExiste("VIANDES") = False Then
            Call Reporting("VIANDES", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("VIANDES")

        NbDenrees = GetNonEmptyRows() - 1

        If NbDenrees > MaxDenrees Then
            Call Reporting("VIANDES", "ARRET", "Nombre de viandes d�passe la dimension > " & MaxDenrees, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        If NbDenrees > 0 Then
            col1 = 2            ' colonne B      Pr�paration du tri des donn�es Viandes
            Mode1 = eSortOrder.Descending       'd�croissant
            col2 = 0
            Mode2 = eSortOrder.Ascending
            col3 = 0
            Mode3 = eSortOrder.Ascending
            Call TriMultiple("VIANDES", col1, Mode1, col2, Mode2, col3, Mode3, 6, NbDenrees + 1)

            PTotViande = 0

            For i = 1 To NbDenrees
                Description(i) = wsExcel.Cells(i + 1, 1).Value
                Call VerifNumeric(wsExcel.Cells(i + 1, 2).Value, "Reel", "VIANDES", i, 2)
                Poids(i) = ValeurR
                'Console.WriteLine("Ligne " & i + 1 & " col 2= " & wsExcel.Cells(i + 1, 2).Value & "  Poids= " & Poids(i) & "  NbErreur= " & NbErreur)
                Call VerifNumeric(wsExcel.Cells(i + 1, 3).Value, "Entier", "VIANDES", i, 3)
                Quant(i) = ValeurI
                'Console.WriteLine("Ligne " & i + 1 & " col 3= " & wsExcel.Cells(i + 1, 3).Value & "  Quant= " & Quant(i) & "  NbErreur= " & NbErreur)
                'StrOption = Console.ReadLine()

                ViandeSC(i) = False                 'SC = Sans Cochon = musulman
                If wsExcel.Cells(i + 1, 4).Value = 1 Then ViandeSC(i) = True
                ViandeSV(i) = False                 'SV = Sans Viande = vegan
                If wsExcel.Cells(i + 1, 5).Value = 1 Then ViandeSV(i) = True
                ResteQuant(i) = Quant(i)            'initialise le Reste � la quantit� initiale
                CPViande = wsExcel.Cells(i + 1, 6).Value
                PTotViande += Poids(i) * Quant(i)       'cumul du poids*quantit� avec le total
            Next
        End If

        '*****************************************************************
        '   Lecture des pr�parations
        '******************************************************************

        If FeuilleExiste("PREPARATIONS") = False Then
            Call Reporting("PREPARATIONS", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("PREPARATIONS")

        NbPreparations = GetNonEmptyRows() - 1

        If NbPreparations > MaxDenrees Then
            Call Reporting("PREPARATIONS", "ARRET", "Nombre de preparations d�passe la dimension > " & MaxDenrees, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        If NbPreparations > 0 Then

            col1 = 2            'colonne B
            Mode1 = eSortOrder.Ascending       'croissant
            col2 = 0
            Mode2 = eSortOrder.Ascending
            col3 = 0
            Mode3 = eSortOrder.Ascending
            Call TriMultiple("PREPARATIONS", col1, Mode1, col2, Mode2, col3, Mode3, 8, NbPreparations + 1)

            PTotPrepa = 0
            QuantTotPrepa = 0

            For i = 1 To NbPreparations
                Preparation(i) = wsExcel.Cells(i + 1, 1).Value
                TestPrepa = wsExcel.Cells(i + 1, 2).Value
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

                Call VerifNumeric(wsExcel.Cells(i + 1, 3).Value, "Entier", "PREPARATIONS", i, 3) 'v�rification de la quantit�
                QuantPrepa(i) = ValeurI
                QuantTotPrepa += QuantPrepa(i)
                PTotPrepa += QuantPrepa(i) * PoidsPrepa(i)

                Call VerifNumeric(wsExcel.Cells(i + 1, 7).Value, "Reel", "PREPARATIONS", i, 7)  'v�rification du Poids Global

                wsExcel.Cells(i + 1, 8).Value = PoidsPrepa(i)
                PrepaSC(i) = False
                If wsExcel.Cells(i + 1, 4).Value = 1 Then PrepaSC(i) = True
                PrepaSV(i) = False
                If wsExcel.Cells(i + 1, 5).Value = 1 Then PrepaSV(i) = True

            Next i
        End If

        '***********************************************************
        '  Lecture des salades
        '***********************************************************

        If FeuilleExiste("SALADES") = False Then
            Call Reporting("SALADES", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("SALADES")

        NbSalades = GetNonEmptyRows() - 1

        If NbSalades > MaxDenrees Then
            Call Reporting("SALADES", "ARRET", "Nombre de salades d�passe la dimension > " & MaxDenrees, "SALADES")
            TestErreur = True
            Exit Sub
        End If

        If NbSalades > 0 Then

            col1 = 2
            Mode1 = eSortOrder.Ascending       ' tri croissant
            col2 = 0
            Mode2 = eSortOrder.Ascending
            col3 = 0
            Mode3 = eSortOrder.Ascending
            Call TriMultiple("SALADES", col1, Mode1, col2, Mode2, col3, Mode3, 8, NbSalades + 1)

            PTotSalad = 0
            QuantTotSalad = 0

            wsExcel.Cells(1, 8).Value = "Eqv Poids"
            For i = 1 To NbSalades
                Salade(i) = wsExcel.Cells(i + 1, 1).Value
                TestPrepa = wsExcel.Cells(i + 1, 2).Value
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

                Call VerifNumeric(wsExcel.Cells(i + 1, 3).Value, "Entier", "SALADES", i, 3) 'v�rification de la quantit�
                QuantSalade(i) = ValeurI
                QuantTotSalad += QuantSalade(i)
                PTotSalad += QuantSalade(i) * PoidsSalade(i)

                Call VerifNumeric(wsExcel.Cells(i + 1, 7).Value, "Reel", "SALADES", i, 7)  'v�rification du Poids Global

                wsExcel.Cells(i + 1, 8).Value = PoidsSalade(i)
                SaladeSC(i) = False
                If wsExcel.Cells(i + 1, 4).Value = 1 Then SaladeSC(i) = True
                SaladeSV(i) = False
                If wsExcel.Cells(i + 1, 5).Value = 1 Then SaladeSV(i) = True

            Next i
        End If

        '*****************************************************
        '  lecture des LAITAGES
        '*****************************************************
        If FeuilleExiste("LAITAGES") = False Then
            Call Reporting("LAITAGES", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("LAITAGES")

        NbLaitages = GetNonEmptyRows() - 1

        If NbLaitages > MaxDenrees Then
            Call Reporting("LAITAGES", "ARRET", "Nombre de laitages d�passe la dimension > " & MaxDenrees, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        If NbLaitages > 0 Then

            col1 = 3
            Mode1 = eSortOrder.Ascending       ' tri croissant
            col2 = 4
            Mode2 = eSortOrder.Descending
            col3 = 2
            Mode3 = eSortOrder.Ascending
            Call TriMultiple("LAITAGES", col1, Mode1, col2, Mode2, col3, Mode3, 6, NbLaitages + 1)

            PtotLait = 0
            PtotZeu = 0

            For i = 1 To NbLaitages
                Laitage(i) = wsExcel.Cells(i + 1, 1).Value
                Call VerifNumeric(wsExcel.Cells(i + 1, 2).Value, "Entier", "LAITAGES", i, 2) 'v�rification de la quantit�
                QuantLait(i) = ValeurI

                TestPrepa = wsExcel.Cells(i + 1, 3).Value
                TestPrepa = TestPrepa.ToUpper()
                CatLait(i) = TestPrepa
                Equiv(i) = wsExcel.Cells(i + 1, 4).Value
                If CatLait(i) = "ZEU" Then
                    PtotZeu += QuantLait(i) * Equiv(i)
                Else
                    PtotLait += QuantLait(i) * Equiv(i)
                End If
                CPLait = wsExcel.Cells(i + 1, 5).Value

                Call VerifNumeric(wsExcel.Cells(i + 1, 6).Value, "Reel", "LAITAGES", i, 6)  'v�rification du Poids Global

            Next i
        End If

        '*****************************************************************
        '   Lecture des anticipes
        '******************************************************************
        Nbanti = 0                   ' nbre de lignes de l'onglet anticip�s
        If FeuilleExiste("ANTICIPES") = True Then

            wsExcel = wbExcel.Worksheets("ANTICIPES")
            Nbanti = GetNonEmptyRows() - 1

            If Nbanti > MaxDenrees Then
                Call Reporting("ANTICIPES", "ARRET", "Nombre de denr�es Anticipees d�passe la dimension > " & MaxDenrees, "RAPPORT")
                TestErreur = True
                Exit Sub
            End If
        Else
            wsExcel = wbExcel.Worksheets.Add("ANTICIPES")
            wsExcel.Name = "ANTICIPES"
        End If
        If Nbanti > 0 Then
            ' tri des anticipes
            col1 = 2
            Mode1 = eSortOrder.Descending
            col2 = 3
            Mode2 = eSortOrder.Ascending
            col3 = 4
            Mode3 = eSortOrder.Descending
            Call TriMultiple("ANTICIPES", col1, Mode1, col2, Mode2, col3, Mode3, 7, Nbanti + 1)

            For i = 1 To Nbanti
                NomAnti(i) = wsExcel.Cells(i + 1, 1).Value          'Nom de la famille
                TestType = wsExcel.Cells(i + 1, 2).Value
                TestType = TestType.Substring(0, 1)
                TestType = TestType.ToUpper()
                If TestType = "V" Or TestType = "P" Or TestType = "S" Then
                    TypeAnti(i) = TestType                          ' type de produit
                Else
                    TexteMsg = "Type " & wsExcel.Cells(i + 1, 2).Value & " � la ligne " & i + 1 & " n'est pas reconnue"
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                    NbErreur += 1
                End If
                DescAnti(i) = wsExcel.Cells(i + 1, 3).Value             'Description produit

                If TypeAnti(i) = "V" Then
                    Call VerifNumeric(wsExcel.Cells(i + 1, 4).Value, "Reel", "ANTICIPES", i, 4)  'v�rification du Poids 
                    PoidsAnti(i) = ValeurR
                Else
                    TestPrepa = wsExcel.Cells(i + 1, 4).Value
                    TestPrepa = TestPrepa.Substring(0, 1)
                    TestPrepa = TestPrepa.ToUpper()
                    TailleAnti(i) = TestPrepa
                    Select Case TestPrepa
                        Case "P"            'Taille petite => poids �quivalent viande = 40 gr
                            PoidsAnti(i) = 10 * CoefPrepa
                        Case "M"            'Taille moyenne => poids = 80 gr
                            PoidsAnti(i) = 20 * CoefPrepa
                        Case "G"            'Taille grande => poids = 120 gr
                            PoidsAnti(i) = 30 * CoefPrepa
                        Case Else
                            TexteMsg = "Preparation ou salade " & DescAnti(i) & " Taille Petit-Moyen-Gros non reconnue"
                            Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                            NbErreur += 1
                    End Select
                End If

                Call VerifNumeric(wsExcel.Cells(i + 1, 5).Value, "Entier", "ANTICIPES", i, 5) 'v�rification de la quantit�
                QuantAnti(i) = ValeurI

                CodePrixanti(i) = wsExcel.Cells(i + 1, 6).Value                             'CodePrix
                Call VerifNumeric(wsExcel.Cells(i + 1, 7).Value, "Reel", "ANTICIPES", i, 7)  'v�rification du Poids Global

                '  Cumul des poids suivant le type de produit
                If TypeAnti(i) = "V" Then PTotViande += PoidsAnti(i) * QuantAnti(i)
                If TypeAnti(i) = "P" Then PTotPrepa += PoidsAnti(i) * QuantAnti(i)
                If TypeAnti(i) = "S" Then PTotSalad += PoidsAnti(i) * QuantAnti(i)

            Next
        End If



        '*********************************************************
        '  lecture des FAMILLES
        '*********************************************************
        If FeuilleExiste("FAMILLES") = False Then
            Call Reporting("FAMILLES", "ARRET", "Feuille manquante", "RAPPORT")
            TestErreur = True
            Exit Sub
        End If
        wsExcel = wbExcel.Worksheets("FAMILLES")
        NbFamilles = GetNonEmptyRows() - 1
        If NbFamilles > MaxFamilles Then
            Call Reporting("FAMILLES", "ARRET", "Nombre de laitages d�passe la dimension > " & MaxFamilles, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        wsExcel.Cells(1, 8).Value = "Test SCSV"
        wsExcel.Cells(1, 10).Value = "Random"

        For i = 1 To NbFamilles
            ' concat�nation du test SC et du test SV en colonne 8 (H) (pour effectuer le tri sur 3 colonnes maxi)
            wsExcel.Cells(i + 1, 8).Value = wsExcel.Cells(i + 1, 7).Value * 10 + wsExcel.Cells(i + 1, 6).Value
            ' attribue un ordre al�atoire � chaque famille 
            wsExcel.Cells(i + 1, 10).Value = A1.Next(NbFamilles)
            ' calcul le nbre de b�n�ficiaires: une part par adulte et une demi-part par enfant, arrondi � l'unit� sup
            wsExcel.Cells(i + 1, 5).Value = Math.Round(wsExcel.Cells(i + 1, 3).Value + 0.51 * wsExcel.Cells(i + 1, 4).Value)
        Next

        col1 = 8                            'colonne H
        Mode1 = eSortOrder.Descending       ' tri croissant
        col2 = 5
        Mode2 = eSortOrder.Descending
        col3 = 10
        Mode3 = eSortOrder.Ascending
        Call TriMultiple("FAMILLES", col1, Mode1, col2, Mode2, col3, Mode3, 10, NbFamilles + 1)

        NbTotViande = 0
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).Value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).Value
            NBenef(i) = wsExcel.Cells(i + 1, 5).Value
            SansCochon(i) = False
            SansViande(i) = False
            If wsExcel.Cells(i + 1, 6).Value = 1 Then SansCochon(i) = True
            If wsExcel.Cells(i + 1, 7).Value = 1 Then SansViande(i) = True
            TestSCSV(i) = wsExcel.Cells(i + 1, 8).Value
            NbTotViande += NBenef(i)
            Panier(i) = 0
        Next
        ' teste le nombre de b�n�ficiaires
        If NbTotViande = 0 Then
            Call Reporting("FAMILLES", "ARRET", " Le nombre de b�n�ficiaires est nul", "FAMILLES")
            TestErreur = True
            Exit Sub
        End If

        ' v�rifie la pr�sence des noms de l'onglet ANTICIPES dans l'onglet Familles
        If Nbanti > 0 Then
            For i = 1 To Nbanti
                TestNom = NomAnti(i)
                Saut = False
                For j = 1 To NbFamilles
                    If TestNom = NomFamille(j) Then
                        Saut = True
                        Exit For
                    End If
                Next
                If Saut = False Then
                    TexteMsg = "Nom de la famille " & TestNom & " � la ligne " & i + 1 & " n'est pas reconnu dans la liste des familles"
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                    NbErreur += 1
                End If
            Next
        End If

        'calcul du poids th�orique par familles = poids total reparti prorata nbre de b�n�ficiaires
        ModuleViande = PTotViande / NbTotViande
        For i = 1 To NbFamilles
            PoidsTheo(i) = ModuleViande * NBenef(i)
        Next

        '**********************************************************
        '   FIN DE LECTURE ,  ARRET SANS AIDA SI ERREURS
        '**********************************************************
        If NbErreur > 0 Then
            TestErreur = True
            Exit Sub
        End If
        '*********************************************************
        ' Mise en forme onglet RESULTATS
        '*********************************************************

        If FeuilleExiste("RESULTATS") = True Then
            wbExcel.Worksheets.Delete("RESULTATS") 'supprime la feuille avant de commencer
        End If
        wsExcel = wbExcel.Worksheets.Add("RESULTATS")                  'ajoute une nouvelle feuille
        wsExcel.Name = "RESULTATS"

        wsExcel.Cells(1, 1).Value = "Nr CAISSE"
        wsExcel.Cells(1, 2).Value = "FAMILLE"
        wsExcel.Cells(1, 3).Value = "Beneficiaires"
        wsExcel.Cells(1, 4).Value = "Sans Cochon"
        wsExcel.Cells(1, 5).Value = "Sans Viande"

        For i = 1 To NbFamilles
            wsExcel.Cells(i + 1, 1).Value = NumFamille(i)
            wsExcel.Cells(i + 1, 2).Value = NomFamille(i)
            wsExcel.Cells(i + 1, 3).Value = NBenef(i)
            If SansCochon(i) Then wsExcel.Cells(i + 1, 4).Value = "OUI"
            If SansViande(i) Then wsExcel.Cells(i + 1, 5).Value = "OUI"
        Next

        Decal = 5

        '***************************************************************
        '   VIANDES  ANTICIPEES
        '***************************************************************

        If Nbanti > 0 Then
            NbDenreesAnti = 0                    ' nbre de colonnes des anticip�s
            DescPrecedent = ""
            PoidsPrecedent = 0
            QTotalAnti = 0
            For i = 1 To Nbanti
                If TypeAnti(i) = "V" Then
                    'd�tecte les lignes cons�cutives de m�me produits
                    If DescAnti(i) <> DescPrecedent Or PoidsAnti(i) <> PoidsPrecedent Then
                        NbDenreesAnti += 1
                        DescPrecedent = DescAnti(i)
                        PoidsPrecedent = PoidsAnti(i)
                        QTotalAnti = 0
                    End If
                    ' en-t�te de ligne
                    QTotalAnti += QuantAnti(i)
                    wsExcel.Cells(1, NbDenreesAnti + Decal).Value = DescAnti(i) & " " & PoidsAnti(i) & " Gr (" & QTotalAnti & ")"

                    For j = 1 To NbFamilles
                        If NomFamille(j) = NomAnti(i) Then
                            wsExcel.Cells(j + 1, NbDenreesAnti + Decal).Value = QuantAnti(i)   'attribue la distribution � la famille
                            Panier(j) += PoidsAnti(i) * QuantAnti(i)                            'incr�mente le panier
                            Exit For
                        End If
                    Next
                End If
            Next

            Decal += NbDenreesAnti

        End If

        If NbDenrees > 0 Then
            For i = 1 To NbDenrees
                wsExcel.Cells(1, i + Decal).Value = Description(i) & " " & Poids(i) & " Gr (" & Quant(i) & ")"
            Next
            wsExcel.Cells(1, NbDenrees + Decal + 1).Value = "Poids attribue"
            wsExcel.Cells(1, NbDenrees + Decal + 2).Value = "Poids theorique"
            wsExcel.Cells(1, NbDenrees + Decal + 3).Value = "Ecart"
            wsExcel.Cells(NbFamilles + 2, 1).Value = "SOMME"
            wsExcel.Cells(NbFamilles + 2, 2).Value = "SOMME"


            '***************************************************************
            '   ATTRIBUTION DES VIANDES
            '***************************************************************
            'priorit� 1: familles sans viande et sans cochon

            Call Attribution1(NbDenrees, ResteQuant, Poids, ViandeSC, ViandeSV)

            Call Attribution2(NbDenrees, ModuleViande, ResteQuant, Poids, ViandeSC, ViandeSV)

            ParamEcart = 1             'type d'�cart = �cart calcul�
            Call Attribution3(NbDenrees, ResteQuant, Poids, ViandeSC, ViandeSV, ParamEcart)
            If TestErreur Then Exit Sub

            'impression des r�sultats

            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, NbDenrees + Decal + 1).Value = Panier(i)
                wsExcel.Cells(i + 1, NbDenrees + Decal + 2).Value = PoidsTheo(i)
                wsExcel.Cells(i + 1, NbDenrees + Decal + 3).Value = Panier(i) - PoidsTheo(i)
            Next

            wsExcel = wbExcel.Worksheets("FAMILLES")
            wsExcel.Cells(1, 9).Value = "ECART"
            ' reporte l'�cart entre la dotation th�orique et r�alis� pour prioriser l'attribution des plats pr�par�s
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, 9).Value = Panier(i) - PoidsTheo(i)
            Next

            Decal += NbDenrees + 3
        End If      'Fin du test s'il n'y a pas de viande

        '*******************************************************************************
        '   TRI DES FAMILLES
        '*******************************************************************************
        ' le tri des onglets FAMILLES et RESULTATS se fait m�me sans plats prepar�s car 
        ' il sert aussi aux salades

        'tri des familles
        wsExcel = wbExcel.Worksheets("FAMILLES")
        col1 = 7            'Col G
        Mode1 = eSortOrder.Descending       ' tri descending
        col2 = 6            'Col F
        Mode2 = eSortOrder.Descending
        If NbDenrees > 0 Then
            ' tri sur les �carts de poids (attribu� - th�orique) pour prioriser l'attribution des pr�parations
            col3 = 9            'col I
            Mode3 = eSortOrder.Ascending
        Else
            col3 = 5
            Mode3 = eSortOrder.Descending
        End If
        Call TriMultiple("FAMILLES", col1, Mode1, col2, Mode2, col3, Mode3, 10, NbFamilles + 1)

        'on relit les familles apr�s le tri
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).Value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).Value
            NBenef(i) = wsExcel.Cells(i + 1, 5).Value
            SansCochon(i) = False
            SansViande(i) = False
            If wsExcel.Cells(i + 1, 6).Value = 1 Then SansCochon(i) = True
            If wsExcel.Cells(i + 1, 7).Value = 1 Then SansViande(i) = True
            TestSCSV(i) = wsExcel.Cells(i + 1, 8).Value
            Panier(i) = wsExcel.Cells(i + 1, 9).Value
        Next

        ' tri des r�sultats, de la m�me fa�on
        wsExcel = wbExcel.Worksheets("RESULTATS")
        col1 = 5
        Mode1 = eSortOrder.Descending       ' tri descending
        col2 = 4
        Mode2 = eSortOrder.Descending
        If NbDenrees > 0 Then
            col3 = Decal
            Mode3 = eSortOrder.Ascending
        Else
            col3 = 3
            Mode3 = eSortOrder.Descending
        End If
        k = NbDenrees + NbPreparations + NbSalades + NbLaitages + Nbanti + 30           'estimation nombre de colonnes
        Call TriMultiple("RESULTATS", col1, Mode1, col2, Mode2, col3, Mode3, k, NbFamilles + 2)

        '***************************************************************************
        '    PLATS PREPARES ANTICIPES
        '***************************************************************************

        If Nbanti > 0 Then

            NbPrepaAnti = 0                    ' nbre de colonnes des anticip�s
            DescPrecedent = ""
            PoidsPrecedent = 0
            QTotalAnti = 0
            For i = 1 To Nbanti
                If TypeAnti(i) = "P" Then
                    'd�tecte les lignes cons�cutives de m�me produits
                    If DescAnti(i) <> DescPrecedent Or PoidsAnti(i) <> PoidsPrecedent Then
                        NbPrepaAnti += 1
                        DescPrecedent = DescAnti(i)
                        PoidsPrecedent = PoidsAnti(i)
                        QTotalAnti = 0
                    End If
                    QTotalAnti += QuantAnti(i)
                    wsExcel.Cells(1, NbPrepaAnti + Decal).Value = DescAnti(i) & " " & TailleAnti(i) & " (" & QTotalAnti & ")"
                    For j = 1 To NbFamilles
                        If NomFamille(j) = NomAnti(i) Then
                            wsExcel.Cells(j + 1, NbPrepaAnti + Decal).Value = QuantAnti(i)
                            Panier(j) += PoidsAnti(i) * QuantAnti(i)
                            Exit For
                        End If
                    Next
                End If
            Next i
            Decal += NbPrepaAnti
        End If

        If NbPreparations > 0 Then

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
                wsExcel.Cells(1, i + Decal).Value = Preparation(i) & " " & TaillePrepa(i) & " (" & QuantPrepa(i) & ")"
            Next
            wsExcel.Cells(1, NbPreparations + Decal + 1).Value = "Nbre attribue"

            '*******************************************************************
            ' Attribution des Plats Prepar�s
            '*******************************************************************
            Call Attribution4(NbPreparations, ResteQuant, PoidsPrepa, PrepaSC, PrepaSV)

            ParamEcart = 2      'ecart donn� par l'ordre des familles, en colonne I
            Call Attribution3(NbPreparations, ResteQuant, PoidsPrepa, PrepaSC, PrepaSV, ParamEcart)
            If TestErreur Then Exit Sub

            'report nbre de plats prepar�s par famille
            SousTotal = NbPrepaAnti + NbPreparations
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, Decal + NbPreparations + 1).FormulaR1C1 = "=SUM(RC[-" & SousTotal & "]:RC[-1])"
            Next

            wsExcel = wbExcel.Worksheets("FAMILLES")
            wsExcel.Cells(1, 9).Value = "Ecart"
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, 9).Value = Panier(i) - PoidsTheo(i)
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")
            Decal = Decal + NbPreparations + 1
        End If



        '***************************************************************************
        '    SALADES ANTICIPES
        '***************************************************************************

        If Nbanti > 0 Then

            NbSaladeAnti = 0                    ' nbre de colonnes des anticip�s
            DescPrecedent = ""
            PoidsPrecedent = 0
            QTotalAnti = 0
            For i = 1 To Nbanti
                If TypeAnti(i) = "S" Then
                    'd�tecte les lignes cons�cutives de m�me produits
                    If DescAnti(i) <> DescPrecedent Or PoidsAnti(i) <> PoidsPrecedent Then
                        NbSaladeAnti += 1
                        DescPrecedent = DescAnti(i)
                        PoidsPrecedent = PoidsAnti(i)
                        QTotalAnti = 0
                    End If
                    QTotalAnti += QuantAnti(i)
                    wsExcel.Cells(1, NbSaladeAnti + Decal).Value = DescAnti(i) & " " & TailleAnti(i) & " (" & QTotalAnti & ")"
                    For j = 1 To NbFamilles
                        If NomFamille(j) = NomAnti(i) Then
                            wsExcel.Cells(j + 1, NbSaladeAnti + Decal).Value = QuantAnti(i)
                            Panier(j) += PoidsAnti(i) * QuantAnti(i)
                            Exit For
                        End If
                    Next
                End If
            Next

            Decal += NbSaladeAnti
        End If

        '************************************************************************
        '    SALADES
        '************************************************************************

        If NbSalades > 0 Then

            wsExcel = wbExcel.Worksheets("FAMILLES")
            For i = 1 To NbFamilles
                Panier(i) = wsExcel.Cells(i + 1, 7).Value
                PoidsTheo(i) = PTotSalad * NBenef(i) / NbTotViande
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")
            QuantTotSalad = 0

            For i = 1 To NbSalades
                ResteQuant(i) = QuantSalade(i)
                wsExcel.Cells(1, i + Decal).Value = Salade(i) & " " & TailleSalade(i) & " (" & QuantSalade(i) & ")"
                QuantTotSalad += QuantSalade(i)
            Next
            wsExcel.Cells(1, Decal + NbSalades + 1).Value = "Nbre Attribue"

            '***********************************************************************
            ' Attribution des SALADES
            '***********************************************************************
            Call Attribution4(NbSalades, ResteQuant, PoidsSalade, SaladeSC, SaladeSV)

            ParamEcart = 2      'ecart donn� par l'ordre des familles, en colonne I
            Call Attribution3(NbSalades, ResteQuant, PoidsSalade, SaladeSC, SaladeSV, ParamEcart)
            If TestErreur Then Exit Sub

            SousTotal = NbSalades + NbSaladeAnti
            For i = 1 To NbFamilles
                wsExcel.Cells(i + 1, Decal + NbSalades + 1).FormulaR1C1 = "=SUM(RC[-" & SousTotal & "]:RC[-1])"
            Next
            Decal += NbSalades + 1
        End If


        If FeuilleExiste("ANTICIPES") = True Then
            wsExcel = wbExcel.Worksheets("ANTICIPES")
            wsExcel.Cells(1, 10).Value = NbDenreesAnti
            wsExcel.Cells(1, 11).Value = NbPrepaAnti
            wsExcel.Cells(1, 12).Value = NbSaladeAnti
        End If

        '*****************************************************
        '  LAITAGES
        '*****************************************************

        'tri des familles
        wsExcel = wbExcel.Worksheets("FAMILLES")
        col1 = 5
        Mode1 = eSortOrder.Descending       ' tri descending
        col2 = 1
        Mode2 = eSortOrder.Ascending
        col3 = 0
        Mode3 = eSortOrder.Ascending

        Call TriMultiple("FAMILLES", col1, Mode1, col2, Mode2, col3, Mode3, 10, NbFamilles + 1)

        'on relit les familles apr�s le tri
        For i = 1 To NbFamilles
            NumFamille(i) = wsExcel.Cells(i + 1, 1).Value
            NomFamille(i) = wsExcel.Cells(i + 1, 2).Value
            NBenef(i) = wsExcel.Cells(i + 1, 5).Value

            Panier(i) = 0
            PanierZeu(i) = 0
        Next

        wsExcel.Cells("H1:J" & MaxFamilles).Clear()

        ' tri des r�sultats, de la m�me fa�on
        wsExcel = wbExcel.Worksheets("RESULTATS")
        col1 = 3
        Mode1 = eSortOrder.Descending       ' tri descending
        col2 = 1
        Mode2 = eSortOrder.Ascending
        col3 = 0
        Mode3 = eSortOrder.Descending

        k = NbDenrees + NbPreparations + NbSalades + NbLaitages + Nbanti + 30       'nombre de colonnes � trier

        Call TriMultiple("RESULTATS", col1, Mode1, col2, Mode2, col3, Mode3, k, NbFamilles + 2)

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

            For i = 1 To NbLaitages
                wsExcel.Cells(1, i + Decal).Value = Laitage(i) & "_EQV" & Equiv(i) & " (" & QuantLait(i) & ")"
            Next
            wsExcel.Cells(1, NbLaitages + Decal + 1).Value = "Nbre Laitages"
            wsExcel.Cells(1, NbLaitages + Decal + 2).Value = "Nbre d'oeufs"

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
                                        wsExcel.Cells(i + 1, Decal + j).Value += 1
                                        Panier(i) += Equiv(j)
                                        ResteQuant(j) -= 1
                                    End If
                                End If
                            Else
                                If Equiv(j) = 1 Then
                                    If ResteQuant(j) > 2 Then
                                        TestPanier = 2
                                        Do While TestPanier < CibleLait
                                            wsExcel.Cells(i + 1, Decal + j).Value += 2
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
                            wsExcel.Cells(i + 1, Decal + j).Value += 1
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
                        wsExcel.Cells(NumeMaxi + 1, Decal + j).Value += 2
                        Panier(NumeMaxi) += Equiv(j) * 2
                        ResteQuant(j) -= 2
                    End While
                End If
            Next j

            '--------repartition des impairs------------------------------
            For j = 1 To NbLaitages
                If ResteQuant(j) > 0 And CatLait(j) <> "BCF" And CatLait(j) <> "ZEU" Then
                    Do While ResteQuant(j) > 0
                        Saut = False
                        For i = 1 To NbFamilles
                            Ecart = Panier(i) - PoidsTheo(i)
                            If Ecart < 0 And Saut = False And NBenef(i) > 0 Then
                                wsExcel.Cells(NumeMaxi + 1, Decal + j).Value += 1
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
                                        wsExcel.Cells(i + 1, Decal + j).Value += 1
                                        PanierZeu(i) += Equiv(j)
                                        ResteQuant(j) -= 1
                                    Else
                                        If Equiv(j) > PoidsTheozeu(i) Then
                                            wsExcel.Cells(i + 1, Decal + j).Value += 1
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
                            wsExcel.Cells(i + 1, Decal + j).Value += 1
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
                        SommeZeu += wsExcel.Cells(i + 1, Decal + j).Value
                    Else
                        SommeLait += wsExcel.Cells(i + 1, Decal + j).Value
                    End If
                Next
                wsExcel.Cells(i + 1, Decal + NbLaitages + 1).Value = SommeLait
                wsExcel.Cells(i + 1, Decal + NbLaitages + 2).Value = PanierZeu(i)
            Next

            Decal += NbLaitages + 2
        End If

        '*********************************************************************
        'Formatage onglet DIVERS
        '**********************************************************************
        If FeuilleExiste("DIVERS") = False Then
            wsExcel = wbExcel.Worksheets.Add("DIVERS")
            wsExcel.Name = "DIVERS"
            NbDivers = 0
        Else
            wsExcel = wbExcel.Worksheets("DIVERS")
            NbDivers = GetNonEmptyRows() - 1
        End If

        If NbDivers > MaxDenrees Then
            Call Reporting("DIVERS", "ARRET", "Nombre de Divers d�passe la dimension > " & MaxDenrees, "RAPPORT")
            TestErreur = True
            Exit Sub
        End If

        If NbDivers > 0 Then

            For i = 1 To NbDivers
                Divers(i) = wsExcel.Cells(i + 1, 1).Value
            Next i

            '-----En-t�te des Divers ----------------------------------
            wsExcel = wbExcel.Worksheets("RESULTATS")
            For i = 1 To NbDivers
                wsExcel.Cells(1, i + Decal).Value = Divers(i)
            Next
            Decal += NbDivers

        End If

        '**************************************************************************
        '   SOUS-TOTAUX PAR DENREES  
        '**************************************************************************
        ' (remplace les formules �crites qui ne sont pas interpr�t�es => cells.value = 0 )

        wsExcel = wbExcel.Worksheets("RESULTATS")

        For j = 6 To Decal
            TotalFamille = 0
            For i = 1 To NbFamilles
                TotalFamille += wsExcel.Cells(i + 1, j).Value
            Next i
            wsExcel.Cells(NbFamilles + 2, j).Value = TotalFamille

        Next j

        '**********************************************************************
        '  formattage onglet RESULTATS
        '**********************************************************************

        'mise en gras colonnes familles
        wsExcel.Columns(1, 5).Style.Font.Bold = True
        wsExcel.Columns(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        wsExcel.Columns(1).Width = 10
        wsExcel.Columns(2).Width = 16
        wsExcel.Columns(3).Width = 4
        For Each c In {3, 4, 5}
            With wsExcel.Column(c)
                .Width = 4
                .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End With
        Next
        With wsExcel.Cells("A1:E1").Style
            .Fill.PatternType = ExcelFillStyle.Solid
            .Fill.BackgroundColor.Indexed = 47
        End With

        'Premi�re ligne

        AlphaColTri = AlphaCol(Decal)
        With wsExcel.Cells("C1:" & AlphaColTri & "1").Style
            .HorizontalAlignment = ExcelHorizontalAlignment.Center
            .TextRotation = 90
            .Font.Bold = True
            .Border.Bottom.Style = ExcelBorderStyle.Medium
        End With

        i = 0
        While i < NbFamilles + 2
            i += 2
            With wsExcel.Cells("A" & i & ":" & AlphaColTri & i).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 26
            End With
        End While

        Decal = 5
        'Viandes anticip�es :  mise en couleur de l'ent�te de colonne

        If NbDenreesAnti > 0 Then
            wsExcel.Columns(Decal + 1, Decal + NbDenreesAnti).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbDenreesAnti).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbDenreesAnti)
            With wsExcel.Cells(AlphaColTri & "1:" & AlphaColTri2 & 1).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 42
            End With
            Decal += NbDenreesAnti
        End If

        ' VIANDES   formattage des colonnes
        If NbDenrees > 0 Then

            wsExcel.Columns(Decal + 1, Decal + NbDenrees).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbDenrees).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            AlphaColTri = AlphaCol(Decal + NbDenrees + 1)
            AlphaColTri2 = AlphaCol(Decal + NbDenrees + 3)
            With wsExcel.Columns(Decal + NbDenrees + 1, Decal + NbDenrees + 3)
                .Width = 8
                .Style.Font.Bold = True
                .Style.Numberformat.Format = "0.0"
            End With
            With wsExcel.Cells(AlphaColTri & 1 & ":" & AlphaColTri2 & NbFamilles + 2).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 43
            End With

            Decal += NbDenrees + 3

        End If

        'Pr�parations anticip�es :  mise en couleur de l'ent�te de colonne
        If NbPrepaAnti > 0 Then
            wsExcel.Columns(Decal + 1, Decal + NbPrepaAnti).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbPrepaAnti).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbPrepaAnti)
            With wsExcel.Cells(AlphaColTri & "1:" & AlphaColTri2 & 1).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 42
            End With
            Decal += NbPrepaAnti
        End If

        ' PREPARATIONS   formattage des colonnes
        If NbPreparations > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbPreparations)
            wsExcel.Columns(Decal + 1, Decal + NbPreparations).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbPreparations).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            AlphaColTri = AlphaCol(Decal + NbPreparations + 1)
            With wsExcel.Columns(Decal + NbPreparations + 1)
                .Width = 6
                .Style.Font.Bold = True
                .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End With
            With wsExcel.Cells(AlphaColTri & 1 & ":" & AlphaColTri & NbFamilles + 2).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 43
            End With

            Decal += NbPreparations + 1
        End If

        'Salades anticip�es :  mise en couleur de l'ent�te de colonne
        If NbSaladeAnti > 0 Then
            wsExcel.Columns(Decal + 1, Decal + NbSaladeAnti).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbSaladeAnti).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbSaladeAnti)
            With wsExcel.Cells(AlphaColTri & "1:" & AlphaColTri2 & 1).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 42
            End With
            Decal += NbSaladeAnti
        End If

        ' SALADES  formattage des colonnes
        If NbSalades > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbSalades)
            wsExcel.Columns(Decal + 1, Decal + NbSalades).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbSalades).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            AlphaColTri = AlphaCol(Decal + NbSalades + 1)
            With wsExcel.Columns(Decal + NbSalades + 1)
                .Width = 6
                .Style.Font.Bold = True
                .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End With
            With wsExcel.Cells(AlphaColTri & 1 & ":" & AlphaColTri & NbFamilles + 2).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 43
            End With

            Decal += NbSalades + 1
        End If

        ' LAITAGES  formattage des colonnes
        If NbLaitages > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbLaitages)
            wsExcel.Columns(Decal + 1, Decal + NbLaitages).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbLaitages).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            AlphaColTri = AlphaCol(Decal + NbLaitages + 1)
            AlphaColTri2 = AlphaCol(Decal + NbLaitages + 2)
            With wsExcel.Columns(Decal + NbLaitages + 1, Decal + NbLaitages + 2)
                .Width = 6
                .Style.Font.Bold = True
                .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            End With
            With wsExcel.Cells(AlphaColTri & 1 & ":" & AlphaColTri2 & NbFamilles + 2).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 43
            End With

            Decal += NbLaitages + 1
        End If

        'DIVERS   Formattage des colonnes
        If NbDivers > 0 Then
            AlphaColTri = AlphaCol(Decal + 1)
            AlphaColTri2 = AlphaCol(Decal + NbDivers + 1)
            wsExcel.Columns(Decal + 1, Decal + NbDivers + 1).Width = 4
            wsExcel.Columns(Decal + 1, Decal + NbDivers + 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

        End If

        ' Traits verticaux
        For i = 1 To Decal + NbDivers + 1
            AlphaColTri = AlphaCol(i)
            With wsExcel.Cells(AlphaColTri & "1:" & AlphaColTri & NbFamilles + 2).Style
                .Border.Left.Style = ExcelBorderStyle.Thin
                .Border.Right.Style = ExcelBorderStyle.Thin
            End With

        Next i

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
                            TestPanier = Poids(j) + Panier(i)
                            If TestPanier < PoidsTheo(i) Then
                                wsExcel.Cells(i + 1, Decal + j).Value += 1
                                Panier(i) += Poids(j)
                                Reste(j) -= 1
                            Else
                                If Poids(j) > PoidsTheo(i) Then
                                    wsExcel.Cells(i + 1, Decal + j).Value += 1
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
        ' Anticip�s => v�rifie que le panier ne d�passe pas le poids th�orique
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
                       (SansCochon(i) = False Or Testsc(j) = True) And NBenef(i) > 0 And Panier(i) < PoidsTheo(i) Then
                        wsExcel.Cells(i + 1, Decal + j).Value += 1
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
                        TexteMsg = "(Sub. Attribution3) Arr�t � la denr�e " & j & " sur " & Nbdenrees & " R�partition incompl�te"
                        Call Reporting("RESULTATS", "ARRET", TexteMsg, "RESULTATS")
                        TestErreur = True
                        Exit Sub
                    End If
                Else
                    NumeMaxi += 1
                    If NumeMaxi > NbFamilles Then NumeMaxi = 1
                End If
                'attribution des denr�es une par une
                If (SansViande(NumeMaxi) = False Or Testsv(j) = True) And
                    (SansCochon(NumeMaxi) = False Or Testsc(j) = True) And NBenef(NumeMaxi) > 0 Then
                    wsExcel.Cells(NumeMaxi + 1, Decal + j).Value += 1
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
                            wsExcel.Cells(i + 1, Decal + j).Value += 1
                            Panier(i) += Poids(j)
                            Reste(j) -= 1
                        End If
                    End If
                Next j
            End If
        Next i

    End Sub

    Public Function GetNonEmptyRows() As Integer
        ' compte le nombre de lignes non vides en colonne 1
        Dim total = 0
        While wsExcel.Cells(total + 1, 1).Value <> Nothing
            total += 1
        End While
        GetNonEmptyRows = total
    End Function

    Public Function FeuilleExiste(FeuilleAVerifier As String) As Boolean
        'fonction qui v�rifie si la "FeuilleAVerifier" existe dans le Classeur actif

        On Error GoTo SiErreur
        Dim Feuille
        FeuilleExiste = False
        For Each Feuille In wbExcel.Worksheets
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

    Public Sub VerifNumeric(Cellule As Object, TypeRetour As String, Onglet As String, i As Integer, col As Integer)
        ' V�rifie si une cellule est bien num�rique:
        ' re�oit la cellule, le type recherch�, et sa localisation en cas d'erreur: onglet, ligne et colonne 
        ' renvoie la donn�e ValeurR ou ValeurI si on veut du type Single ou Integer  (variables publiques)
        ' ne fait rien si cellule vide
        ' sinon affiche un message dans l'onglet Reporting avec l'onglet, la ligne et la colonne
        ' impose une sortie avec TestErreur = true car sinon, bug � l'utilisation  (variable publique)
        '--------------------------------------------------------------------------------------------
        'Dim stroption As String
        Dim ColAlpha As String              ' conversion de la colonne en alphanum�rique

        ValeurR = 0
        ValeurI = 0
        If IsNumeric(Cellule) Then
            If TypeRetour = "Reel" Then ValeurR = CType(Cellule, Single)
            If TypeRetour = "Entier" Then ValeurI = CType(Cellule, Integer)

        ElseIf Cellule <> "" Then
            ColAlpha = AlphaCol(col)
            TexteMsg = " A la ligne " & i + 1 & " Colonne " & ColAlpha & ", contenu:  " & Cellule & " n'est pas un nombre!"
            Call Reporting(Onglet, "ARRET", TexteMsg, Onglet)
            NbErreur += 1
            TestErreur = True
        End If
    End Sub
    Public Sub TriMultiple(Feuille As String, Col1 As Integer, Mode1 As eSortOrder, Col2 As Integer, Mode2 As eSortOrder,
                    Col3 As Integer, Mode3 As eSortOrder, nbcol As Integer, nblignes As Integer)
        '------------------------------------------------------------------------------------
        ' Il y a toujours une ligne d'ent�tes

        Dim AlphaColTri As String

        wsExcel = wbExcel.Worksheets(Feuille)
        AlphaColTri = AlphaCol(nbcol)
        Col1 -= 1                   'Enl�ve -1 car les num�ros de colonne d�butent � 0: col 0 = col A
        Col2 -= 1
        Col3 -= 1

        Dim sortOptions = RangeSortOptions.Create()
        sortOptions.CompareOptions = CompareOptions.OrdinalIgnoreCase
        Dim sortOptionsC1 = sortOptions.SortBy.Column(Col1, Mode1)
        Dim sortOptionsC2 = If(Col2 >= 0, sortOptionsC1.ThenSortBy.Column(Col2, Mode2), sortOptionsC1)
        Dim sortOptionsC3 = If(Col3 >= 0, sortOptionsC2.ThenSortBy.Column(Col3, Mode3), sortOptionsC2)

        wsExcel.Cells("A2:" & AlphaColTri & nblignes).Sort(sortOptions)
    End Sub


    Public Sub Reporting(Onglet As String, Criticite As String, ReportMsg As String, Retour As String)

        wsExcel = wbExcel.Worksheets("RAPPORT")
        nbReport += 1
        wsExcel.Cells(nbReport, 1).Value = Onglet
        wsExcel.Cells(nbReport, 2).Value = ReportMsg
        wsExcel.Cells(nbReport, 8).Value = Criticite
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
        Dim NbErreur As Integer
        Dim EnTete As String
        Dim Complet As String
        Dim Separ As String = "Gr ("
        Dim IntPos As Integer
        Dim Brut As String
        Dim TotalFamille As Single

        '--------D�claration variables Anticipes
        Dim TypeAnti(MaxDenrees) As String
        Dim DescAnti(MaxDenrees) As String
        Dim PoidsAnti(MaxDenrees) As Single
        Dim TypeReduit(MaxDenrees) As String
        Dim CodeReduit(MaxDenrees) As String
        Dim PoidsReduit(MaxDenrees) As Single
        Dim NbAnti As Integer
        Dim NbDenreesAnti As Integer
        Dim NbPrepaAnti As Integer
        Dim NbSaladeAnti As Integer
        Dim NbDivers As Integer
        Dim TestType As String
        Dim TestAnti As Boolean

        TestErreur = False              ' initialisation du Test sur Nbre d'erreurs
        NbErreur = 0

        wsExcel = wbExcel.Worksheets("FAMILLES")
        NbFamilles = GetNonEmptyRows() - 1

        wsExcel = wbExcel.Worksheets("ANTICIPES")
        NbAnti = GetNonEmptyRows() - 1

        NbDenreesAnti = wsExcel.Cells(1, 10).Value          'relecture du nombre de viandes anticip�es
        NbPrepaAnti = wsExcel.Cells(1, 11).Value            'relecture du nombre de preparations anticip�es
        NbSaladeAnti = wsExcel.Cells(1, 12).Value           'relecture du nombre de salades anticip�es

        '----------relecture des donn�es anticip�es Description, type et poids ---------------------------------------------------
        If NbAnti > 0 Then

            For i = 1 To NbAnti
                TestType = wsExcel.Cells(i + 1, 2).Value            ' d�code le type de denr�e anticip�e
                TestType = TestType.Substring(0, 1)
                TestType = TestType.ToUpper()
                If TestType = "V" Or TestType = "P" Or TestType = "S" Then
                    TypeAnti(i) = TestType                          ' type de produit
                Else
                    TexteMsg = "Type " & wsExcel.Cells(i + 1, 2).Value & " � la ligne " & i + 1 & " n'est pas reconnue"
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                    NbErreur += 1
                End If
                DescAnti(i) = wsExcel.Cells(i + 1, 3).Value             'Description produit

                If TypeAnti(i) = "V" Then
                    Call VerifNumeric(wsExcel.Cells(i + 1, 4).Value, "Reel", "ANTICIPES", i, 4)
                    If TestErreur Then Exit Sub
                    PoidsAnti(i) = ValeurR
                End If

            Next i
        End If

        Decal = 5

        If NbDenreesAnti > 0 Then
            '------calcul du panier pour viandes anticip�es -----------------------------------------------------
            wsExcel = wbExcel.Worksheets("RESULTATS")

            For k = 1 To NbDenreesAnti
                Complet = wsExcel.Cells(1, k + Decal).Value             ' reprend l'ent�te de colonne
                IntPos = InStr(Complet, Separ)
                Brut = Complet.Remove(IntPos - 2)                   ' ent�te r�duite � description & poids 
                TestAnti = False

                For j = 1 To NbAnti                                 ' recherche dans la liste des anticip�s la denr�e pour trouver le poids
                    EnTete = DescAnti(j) & " " & PoidsAnti(j)       ' calcule l'ent�te r�duite
                    If Brut = EnTete Then
                        Poids(k) = PoidsAnti(j)
                        TestAnti = True ' si en-t�te ok, affecte le poids � la denr�e de la colonne
                        Exit For
                    End If
                Next j
                If TestAnti = False Then
                    TexteMsg = "Denr�e en Colonne " & k + Decal & ":  " & Complet & " non reconue dans l'onglet ANTICIPES, pas de synth�se AIDA"
                    Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
                End If

                TotalFamille = 0
                For i = 1 To NbFamilles                                 ' balaye les familles avec les donn�es de la denr�e en t�te
                    Call VerifNumeric(wsExcel.Cells(i + 1, Decal + k).Value, "Entier", "RESULTATS", i, Decal + k)
                    If TestErreur Then Exit Sub
                    Panier(i) += ValeurI * Poids(k)             ' incr�mente le panier = quantit� *poids
                    TotalFamille += ValeurI
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + k).Value = TotalFamille
            Next k
        End If

        Decal += NbDenreesAnti

        '---------calcul du panier pour les viandes ------------------------------------------------------
        wsExcel = wbExcel.Worksheets("VIANDES")
        NbDenrees = GetNonEmptyRows() - 1

        If NbDenrees > 0 Then
            For j = 1 To NbDenrees                      'retrouve le poids des denr�es
                Poids(j) = wsExcel.Cells(j + 1, 2).Value
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")
            For i = 1 To NbFamilles
                For j = 1 To NbDenrees
                    Call VerifNumeric(wsExcel.Cells(i + 1, Decal + j).Value, "Entier", "RESULTATS", i, Decal + j)
                    If TestErreur Then Exit Sub
                    Panier(i) += ValeurI * Poids(j)             ' incr�mente le panier = quantit� *poids
                Next
                '  -------- �criture du panier apr�s mise � jour -------------------------------------------------
                wsExcel.Cells(i + 1, Decal + NbDenrees + 1).Value = Panier(i)
                wsExcel.Cells(i + 1, Decal + NbDenrees + 3).Value = Panier(i) - wsExcel.Cells(i + 1, Decal + NbDenrees + 2).Value
            Next

            '---------calcul des totaux par produit--------------------------
            For j = 1 To NbDenrees
                TotalFamille = 0
                For i = 1 To NbFamilles
                    TotalFamille += wsExcel.Cells(i + 1, Decal + j).Value
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + j).Value = TotalFamille
            Next j

            Decal += NbDenrees + 3

        End If

        '------------PREPARATIONS --------------------------------------------------------------------
        wsExcel = wbExcel.Worksheets("PREPARATIONS")
        NbPreparations = GetNonEmptyRows() - 1

        If NbPreparations > 0 Then
            wsExcel = wbExcel.Worksheets("RESULTATS")
            '---------calcul des totaux par produit--------------------------
            For j = 1 To NbPreparations
                TotalFamille = 0
                For i = 1 To NbFamilles
                    TotalFamille += wsExcel.Cells(i + 1, Decal + j).Value
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + j).Value = TotalFamille
            Next j
            Decal += NbPreparations + NbPrepaAnti + 1
        End If

        '-------------------SALADES --------------------------------------------
        wsExcel = wbExcel.Worksheets("SALADES")
        NbSalades = GetNonEmptyRows() - 1

        If NbSalades > 0 Then
            wsExcel = wbExcel.Worksheets("RESULTATS")
            '---------calcul des totaux par produit--------------------------
            For j = 1 To NbSalades
                TotalFamille = 0
                For i = 1 To NbFamilles
                    TotalFamille += wsExcel.Cells(i + 1, Decal + j).Value
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + j).Value = TotalFamille
            Next j
            Decal += NbSalades + NbSaladeAnti + 1
        End If

        '------calcul des totaux pour les laitages ----------------------------------------------------------
        wsExcel = wbExcel.Worksheets("LAITAGES")
        NbLaitages = GetNonEmptyRows() - 1
        If NbLaitages > 0 Then
            For i = 1 To NbLaitages
                CatLait(i) = (wsExcel.Cells(i + 1, 3).Value).toupper()
                Equiv(i) = wsExcel.Cells(i + 1, 4).Value
            Next

            wsExcel = wbExcel.Worksheets("RESULTATS")

            For i = 1 To NbFamilles
                SommeLait = 0
                SommeZeu = 0
                For j = 1 To NbLaitages
                    Call VerifNumeric(wsExcel.Cells(i + 1, Decal + j).Value, "Entier", "RESULTATS", i, Decal + j)
                    If TestErreur Then Exit Sub
                    If CatLait(j) = "ZEU" Then
                        SommeZeu += ValeurI * Equiv(j)
                    Else
                        SommeLait += ValeurI
                    End If
                Next
                wsExcel.Cells(i + 1, Decal + NbLaitages + 1).Value = SommeLait
                wsExcel.Cells(i + 1, Decal + NbLaitages + 2).Value = SommeZeu
            Next
            '---------calcul des totaux par produit--------------------------
            For j = 1 To NbLaitages
                TotalFamille = 0
                For i = 1 To NbFamilles
                    TotalFamille += wsExcel.Cells(i + 1, Decal + j).Value
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + j).Value = TotalFamille
            Next j

            Decal += NbLaitages + 2
        End If

        '------------Sous-totaux Divers ------------------------------------------------------
        wsExcel = wbExcel.Worksheets("DIVERS")
        NbDivers = GetNonEmptyRows() - 1
        If NbDivers > 0 Then
            wsExcel = wbExcel.Worksheets("RESULTATS")
            For j = 1 To NbDivers
                TotalFamille = 0
                For i = 1 To NbFamilles
                    TotalFamille += wsExcel.Cells(i + 1, Decal + j).Value
                Next i
                wsExcel.Cells(NbFamilles + 2, Decal + j).Value = TotalFamille
            Next j
        End If

        ' --------Gestion des erreurs: sortie sans AIDA en cas d'erreurs ------------------------
        If NbErreur > 0 Then TestErreur = True

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
        Dim Ecart As Single
        Dim SousTotal As Single
        Dim TotalArrondi As Single
        Dim Arrondi As Single
        Dim FormatCol As String

        '--------D�claration variables Anticipes
        'Dim NomAnti(MaxDenrees) As String
        Dim TypeAnti(MaxDenrees) As String
        Dim DescAnti(MaxDenrees) As String
        Dim PoidsGlobalAnti(MaxDenrees) As String
        Dim PoidsAnti(MaxDenrees) As Single
        Dim QuantAnti(MaxDenrees) As Integer
        Dim CodePrixAnti(MaxDenrees) As String
        Dim TypeReduit(MaxDenrees) As String
        Dim CodeReduit(MaxDenrees) As String
        Dim PoidsReduit(MaxDenrees) As Single
        Dim PoidsGlobalReduit(MaxDenrees) As Single
        Dim Nbanti As Integer
        Dim TestType As String
        Dim NbSingleAnti As Integer
        Dim NbDenreesAnti As Integer
        Dim NbPrepaAnti As Integer
        Dim NbSaladeAnti As Integer
        Dim DescPrecedent As String
        Dim PoidsPrecedent As Single
        Dim TestPrepa As String
        'Dim StrOption As String
        Dim NbErreur As Integer
        Dim TotalFamille As Single
        Dim PoidsVerif As Single
        Dim Onglet As String

        ' chemin r�seau g�n�rique pour l'enregistrement des images de code-barre
        ' CheminBureau = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        CheminBureau = Environment.CurrentDirectory
        NbErreur = 0

        wsExcel = wbExcel.Worksheets("FAMILLES")
        NbFamilles = GetNonEmptyRows() - 1

        wsExcel = wbExcel.Worksheets("PRIX")
        nbPrix = GetNonEmptyRows() - 1

        '----------------------lecture des codes prix ---------------------------------------
        If nbPrix > MaxPrix Then
            TexteMsg = "Nombre de prix d�passe la dimension > " & MaxPrix
            Call Reporting("PRIX", "ARRET", TexteMsg, "PRIX")
            Exit Sub
        End If

        If nbPrix > 0 Then
            For i = 1 To nbPrix
                CodePrix(i) = wsExcel.Cells(i + 1, 2).Value
                CodeAIDA(i) = wsExcel.Cells(i + 1, 3).Value
                UnitAIDA(i) = wsExcel.Cells(i + 1, 4).Value
                PrixAIDA(i) = wsExcel.Cells(i + 1, 5).Value
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
            Exit Sub
        End If

        NbCat = 0

        ' **** CONSTRUCTION DE LA LISTE UNIQUE et SANS DOUBLONS DES CODES PRIX *****
        ' relecture des onglets de donn�es
        ' v�rification du poids global si unit� kgM
        ' construction de la liste unique des cat�gories
        '-----------------liste des codes viandes------------------------------------
        wsExcel = wbExcel.Worksheets("VIANDES")
        nbdenrees = GetNonEmptyRows() - 1

        If nbdenrees > 0 Then
            For j = 1 To nbdenrees
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value
            Next

            Call ListeCategorie(nbdenrees, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
        UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes plats prepares---------------------------------------
        wsExcel = wbExcel.Worksheets("PREPARATIONS")
        NbPreparations = GetNonEmptyRows() - 1

        If NbPreparations > 0 Then
            For j = 1 To NbPreparations
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value
                PoidsVerif = wsExcel.Cells(j + 1, 7).Value
                TestPrepa = wsExcel.Cells(j + 1, 1).Value
                Onglet = "PREPARATIONS"
                Call VerifPoids(j, nbPrix, CodePrixDenree(j), CodePrix, UnitAIDA, PoidsVerif, TestPrepa, Onglet)
            Next j

            Call ListeCategorie(NbPreparations, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes salades   --------------------------------------
        wsExcel = wbExcel.Worksheets("SALADES")
        NbSalades = GetNonEmptyRows() - 1

        If NbSalades > 0 Then
            For j = 1 To NbSalades
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value
                PoidsVerif = wsExcel.Cells(j + 1, 7).Value
                TestPrepa = wsExcel.Cells(j + 1, 1).Value
                Onglet = "SALADES"
                Call VerifPoids(j, nbPrix, CodePrixDenree(j), CodePrix, UnitAIDA, PoidsVerif, TestPrepa, Onglet)
            Next j

            Call ListeCategorie(NbSalades, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes laitages---------------------------------------
        wsExcel = wbExcel.Worksheets("LAITAGES")
        NbLaitages = GetNonEmptyRows() - 1

        If NbLaitages > 0 Then
            For j = 1 To NbLaitages
                TestPrepa = wsExcel.Cells(j + 1, 1).Value
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 5).Value
                PoidsVerif = wsExcel.Cells(j + 1, 6).Value
                Onglet = "LAITAGES"
                Call VerifPoids(j, nbPrix, CodePrixDenree(j), CodePrix, UnitAIDA, PoidsVerif, TestPrepa, Onglet)
            Next j

            Call ListeCategorie(NbLaitages, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------liste des codes divers---------------------------------------
        wsExcel = wbExcel.Worksheets("DIVERS")
        NbDivers = GetNonEmptyRows() - 1

        If NbDivers > 0 Then
            For j = 1 To NbDivers
                TestPrepa = wsExcel.Cells(j + 1, 1).Value
                CodePrixDenree(j) = wsExcel.Cells(j + 1, 2).Value
                PoidsVerif = wsExcel.Cells(j + 1, 3).Value
                Onglet = "DIVERS"
                Call VerifPoids(j, nbPrix, CodePrixDenree(j), CodePrix, UnitAIDA, PoidsVerif, TestPrepa, Onglet)
            Next j

            Call ListeCategorie(NbDivers, nbPrix, CodePrixDenree, CodePrix, CodeAIDA,
            UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        '------------Relecture des anticipes ----------------------------------------------------
        wsExcel = wbExcel.Worksheets("ANTICIPES")
        Nbanti = GetNonEmptyRows() - 1

        NbDenreesAnti = wsExcel.Cells(1, 10).Value          'relecture du nombre de viandes anticip�es
        NbPrepaAnti = wsExcel.Cells(1, 11).Value            'relecture du nombre de preparations anticip�es
        NbSaladeAnti = wsExcel.Cells(1, 12).Value           'relecture du nombre de salades anticip�es

        If Nbanti > 0 Then
            For i = 1 To Nbanti
                TestType = wsExcel.Cells(i + 1, 2).Value            ' d�code le type de denr�e anticip�e
                TestType = TestType.Substring(0, 1)
                TestType = TestType.ToUpper()
                If TestType = "V" Or TestType = "P" Or TestType = "S" Then
                    TypeAnti(i) = TestType                          ' type de produit
                Else
                    TexteMsg = "Type " & wsExcel.Cells(i + 1, 2).Value & " � la ligne " & i + 1 & " n'est pas reconnue"
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                    NbErreur += 1
                End If
                DescAnti(i) = wsExcel.Cells(i + 1, 3).Value             'Description produit

                If TypeAnti(i) = "V" Then                               'Poids des viandes
                    Call VerifNumeric(wsExcel.Cells(i + 1, 4).Value, "Reel", "ANTICIPES", i, 4)
                    If TestErreur Then Exit Sub
                    PoidsAnti(i) = ValeurR
                Else
                    TestPrepa = wsExcel.Cells(i + 1, 4).Value       'd�code la taille 
                    TestPrepa = TestPrepa.Substring(0, 1)
                    TestPrepa = TestPrepa.ToUpper()

                    Select Case TestPrepa                           ' calcule le poids en fonction de la taille
                        Case "P"                                    'Taille petite => poids �quivalent viande = 40 gr
                            PoidsAnti(i) = 10 * CoefPrepa           'rappel coefPrepa est une constante
                        Case "M"                                    'Taille moyenne => poids = 80 gr
                            PoidsAnti(i) = 20 * CoefPrepa
                        Case "G"                                    'Taille grande => poids = 120 gr
                            PoidsAnti(i) = 30 * CoefPrepa
                        Case Else
                            TexteMsg = "Preparation ou salade, ligne " & i + 1 & " Taille " & DescAnti(i) & " non reconnue"
                            Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "ANTICIPES")
                            NbErreur += 1
                    End Select
                End If

                'v�rification poids global si unit� kgM
                CodePrixAnti(i) = wsExcel.Cells(i + 1, 6).Value           'CodePrix
                PoidsGlobalAnti(i) = wsExcel.Cells(i + 1, 7).Value           'PoidsGlobal pour unit�s KgM
                PoidsVerif = PoidsGlobalAnti(i)
                TestPrepa = wsExcel.Cells(i + 1, 3).Value
                Onglet = "ANTICIPES"
                Call VerifPoids(i, nbPrix, CodePrixAnti(i), CodePrix, UnitAIDA, PoidsVerif, TestPrepa, Onglet)

            Next i
            '-----------liste des codes anticipes------------------------------------------
            Call ListeCategorie(Nbanti, nbPrix, CodePrixAnti, CodePrix, CodeAIDA,
                UnitAIDA, Categorie, CatAIDA, UnAIDA, PrixAIDA, PrixListe)
        End If

        ' -----------------V�rification existence de chaque Code Prix dans la base prix-------------------

        For i = 1 To NbCat
            Test = False
            For j = 1 To nbPrix
                If Categorie(i) = CodePrix(j) Then
                    Test = True
                    Exit For
                End If
            Next j
            If Test = False Then
                TexteMsg = "Code Prix utilis�: " & Categorie(i) & " non d�clar� dans la base PRIX"
                Call Reporting("PRIX", "ALERTE", TexteMsg, "PRIX")
                NbErreur += 1
            End If

            'Console.WriteLine("Categorie " & i & " Code Prix = " & Categorie(i) & " unite " & UnAIDA(i))

        Next i
        'StrOption = Console.ReadLine()
        '-----------Formattage feuille AIDA-----------------------------------------------------
        If FeuilleExiste("AIDA") = True Then
            wbExcel.Worksheets.Delete("AIDA")
        End If

        wsExcel = wbExcel.Worksheets.Add("AIDA")
        wsExcel.Name = "AIDA"

        '--------recopie colonnes familles ---------------------------------------------------
        wsExcel = wbExcel.Worksheets("RESULTATS")
        wsExcel.Cells("A:C").Copy(wbExcel.Worksheets("AIDA").Cells("A:C"))

        wsExcel = wbExcel.Worksheets("AIDA")
        wsExcel.InsertRow(1, 1)
        wsExcel.Row(1).Height = 164

        '---------d�cale l'ent�te des 3 premieres colonnes -----------------------------
        wsExcel.Cells("A2:C2").Copy(wsExcel.Cells("A1:C1"))
        wsExcel.Cells("A2:C2").Clear()

        wsExcel.Columns(1).Width = 10
        wsExcel.Columns(2).Width = 16
        wsExcel.Columns(3).Width = 4
        wsExcel.Columns(4, NbCat + 4).Width = 13

        wsExcel.Rows(2).Height = 15

        AlphaColTri = AlphaCol(NbCat + 4)
        With wsExcel.Cells("D:" & AlphaColTri)
            .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        End With
        wsExcel.Cells(2, NbCat + 4).Value = "PRIX TOTAL"

        Decal = 3
        For j = 1 To NbCat
            wsExcel.Cells(2, Decal + j).Value = Categorie(j) & " " & UnAIDA(j)
        Next

        '------Arret si erreur d�tect�e ---------------------
        If NbErreur > 0 Then Exit Sub

        '------------------V�rification des totaux en bas de colonne---------------------------------
        wsExcel = wbExcel.Worksheets("RESULTATS")
        Decal = 5
        If nbdenrees > 0 Or NbDenreesAnti > 0 Then
            Call TestSomme2(nbdenrees + NbDenreesAnti)      ' teste si somme des quant. attribu�es = quant. d�clar�e et si somme non nulle
            Decal += NbDenreesAnti
            If nbdenrees > 0 Then Decal += nbdenrees + 3
        End If
        If NbPreparations > 0 Or NbPrepaAnti > 0 Then
            Call TestSomme2(NbPreparations + NbPrepaAnti)
            Decal += NbPrepaAnti
            If NbPreparations > 0 Then Decal += NbPreparations + 1
        End If
        If NbSalades > 0 Or NbSaladeAnti > 0 Then
            Call TestSomme2(NbSalades + NbSaladeAnti)
            Decal += NbSaladeAnti
            If NbSalades > 0 Then Decal += NbSalades + 1
        End If
        If NbLaitages > 0 Then
            Call TestSomme2(NbLaitages)
            Decal += NbLaitages + 2
        End If
        If NbDivers > 0 Then
            Call TestSomme(NbDivers)        'pas de quantit� d�clar�e, test si somme nulle seulement
        End If

        '----- Construction de la liste des anticip�s SANS DOUBLONS = liste r�duite ----------------------
        If Nbanti > 0 Then
            NbSingleAnti = 0                    ' nbre de colonnes des anticip�s
            DescPrecedent = ""
            PoidsPrecedent = 0

            For i = 1 To Nbanti
                If DescAnti(i) <> DescPrecedent Or PoidsAnti(i) <> PoidsPrecedent Then      'detecte que ce n'est pas un doublon
                    ' un doublon c'est le m�me produit au m�me poids pour diff�rentes familles
                    NbSingleAnti += 1
                    DescPrecedent = DescAnti(i)
                    PoidsPrecedent = PoidsAnti(i)

                    TypeReduit(NbSingleAnti) = TypeAnti(i)
                    CodeReduit(NbSingleAnti) = CodePrixAnti(i)
                    PoidsReduit(NbSingleAnti) = PoidsAnti(i)
                    PoidsGlobalReduit(NbSingleAnti) = PoidsGlobalAnti(i)

                End If
            Next i
        End If


        ' ****************************************************************************************
        '      CALCUL DES CUMUL  de poids ou de taille, famille par famille
        '******************************************************************************************
        wsExcel = wbExcel.Worksheets("RESULTATS")
        For i = 1 To NbFamilles
            For k = 1 To NbCat
                PoidsCat(k) = 0
            Next k
            Decal = 5

            '--------------Report des cumuls pour les viandes anticip�es --------------------------------------------
            If Nbanti > 0 Then
                k = 0
                For j = 1 To NbSingleAnti
                    If TypeReduit(j) = "V" Then
                        ' dans la liste des anticip�s Sans Doublons, on s�lectionne les viandes
                        k += 1
                        CodePrixDenree(k) = CodeReduit(j)               'Code Prix
                        Poids(k) = PoidsReduit(j) / 1000                ' Poids (valable pour unit� kgC )
                    End If
                Next j

                ' v�rifie que les indices de viandes anticip�es sont justes:
                ' NbDenreesAnti = nombre de viandes anticip�es relues dans ANTICIPES
                ' k= nombre de viandes anticip�es recalcul�es dans la liste r�duite

                If NbDenreesAnti <> k Then
                    TexteMsg = "Nombre de viandes anticip�es inexactes: NbDenreesAnti = " & NbDenreesAnti & "  k = " & k
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
                End If
                Call ReportCumul(NbDenreesAnti, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal += NbDenreesAnti
            End If

            '----------report des cumuls pour les viandes -----------------------------------------------
            If nbdenrees > 0 Then
                wsExcel = wbExcel.Worksheets("VIANDES")
                For j = 1 To nbdenrees
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value
                    Poids(j) = wsExcel.Cells(j + 1, 2).Value / 1000
                Next j
                Call ReportCumul(nbdenrees, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal = Decal + nbdenrees + 3
            End If

            '--------------Report des cumuls pour les preparations anticip�es --------------------------------------------
            If Nbanti > 0 Then
                k = 0
                For j = 1 To NbSingleAnti
                    If TypeReduit(j) = "P" Then
                        ' dans la liste des anticip�s Sans Doublons, on s�lectionne les pr�parations
                        k += 1
                        CodePrixDenree(k) = CodeReduit(j)
                        Poids(k) = PoidsGlobalReduit(j)     'Poids Global, valable pour unit� kgM
                    End If
                Next j
                If NbPrepaAnti <> k Then
                    TexteMsg = "Nombre de pr�parations anticip�es inexactes: nbPrepaAnti = " & NbPrepaAnti & "  k = " & k
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
                End If
                Call ReportCumul(NbPrepaAnti, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal += NbPrepaAnti
            End If
            '--------report des cumuls pour les Pr�parations------------------------------------------------------
            If NbPreparations > 0 Then
                wsExcel = wbExcel.Worksheets("PREPARATIONS")
                For j = 1 To NbPreparations
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value   ' Code Prix
                    Poids(j) = wsExcel.Cells(j + 1, 7).Value
                Next j
                Call ReportCumul(NbPreparations, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal = Decal + NbPreparations + 1
            End If
            '--------------Report des cumuls pour les salades anticip�es --------------------------------------------
            If Nbanti > 0 Then
                k = 0
                For j = 1 To NbSingleAnti
                    If TypeReduit(j) = "S" Then
                        ' dans la liste des anticip�s Sans Doublons, on s�lectionne les pr�parations
                        k += 1
                        CodePrixDenree(k) = CodeReduit(j)       ' Code Prix
                        Poids(k) = PoidsGlobalReduit(j)         'Poids Global, valable pour unit� kgM
                    End If
                Next j
                If NbSaladeAnti <> k Then
                    TexteMsg = "Nombre de salades anticip�es inexactes: nbSaladeAnti = " & NbSaladeAnti & "  k = " & k
                    Call Reporting("ANTICIPES", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
                End If
                Call ReportCumul(NbSaladeAnti, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal += NbSaladeAnti
            End If
            '-----------Report des cumuls pour les salades-----------------------------------------------------
            If NbSalades > 0 Then
                wsExcel = wbExcel.Worksheets("SALADES")
                For j = 1 To NbSalades
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 6).Value
                    Poids(j) = wsExcel.Cells(j + 1, 7).Value
                Next j
                Call ReportCumul(NbSalades, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal = Decal + NbSalades + 1
            End If
            '---------------------------------------------------------------------
            If NbLaitages > 0 Then
                wsExcel = wbExcel.Worksheets("LAITAGES")
                For j = 1 To NbLaitages
                    Equiv(j) = wsExcel.Cells(j + 1, 4).Value
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 5).Value
                    Poids(j) = wsExcel.Cells(j + 1, 6).Value
                Next j
                Call ReportCumul(NbLaitages, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)
                Decal = Decal + NbLaitages + 2
            End If
            '----------------------------------------------------------------------------
            If NbDivers > 0 Then
                wsExcel = wbExcel.Worksheets("DIVERS")
                For j = 1 To NbDivers
                    CodePrixDenree(j) = wsExcel.Cells(j + 1, 2).Value
                    Poids(j) = wsExcel.Cells(j + 1, 3).Value
                Next j
                Call ReportCumul(NbDivers, Decal, i, CodePrixDenree, Categorie, PoidsCat, Poids, UnAIDA, Equiv)

            End If
            '----------------------------------------------------------------------------
            wsExcel = wbExcel.Worksheets("AIDA")        'Ecriture des r�sultats dans l'onglet AIDA
            For k = 1 To NbCat
                wsExcel.Cells(i + 2, k + 3).Value = PoidsCat(k)
            Next k
        Next i              ' fin de boucle sur les familles

        '------------Ajustement des arrondis pour l'unit� KgM-------------------------------
        wsExcel = wbExcel.Worksheets("AIDA")
        For k = 1 To NbCat
            If UnAIDA(k) = "[kgM]" Then
                SousTotal = 0
                TotalArrondi = 0
                For i = 1 To NbFamilles
                    SousTotal += wsExcel.Cells(i + 2, k + 3).Value               'calcule le total brut 
                    Arrondi = Math.Round(wsExcel.Cells(i + 2, k + 3).Value, 2)   ' carrondi les valeurs
                    wsExcel.Cells(i + 2, k + 3).Value = Arrondi
                    TotalArrondi += wsExcel.Cells(i + 2, k + 3).Value           ' calcule le total des arrondis
                Next i
                For i = 1 To NbFamilles
                    If wsExcel.Cells(i + 2, k + 3).Value > 0 Then
                        Ecart = SousTotal - TotalArrondi            ' calcule la diff�rence entre les deux sous-totaux
                        wsExcel.Cells(i + 2, k + 3).Value += Ecart  ' reporte l'�cart entre le total brut et le total des arrondis
                        Exit For            ' sort de la boucle d�s que l'�cart est report�
                    End If
                Next i
            End If
        Next k

        '-----------------------Prix du panier -------------------------------------------
        For i = 1 To NbFamilles
            PrixPanier = 0
            For k = 1 To NbCat
                PrixPanier += wsExcel.Cells(i + 2, k + 3).Value * PrixListe(k)
            Next k
            wsExcel.Cells(i + 2, NbCat + 4).Value = PrixPanier
        Next i

        '------------------sous-totaux par cat�gorie ----------------------------------------

        For j = 1 To NbCat + 1
            TotalFamille = 0
            For i = 1 To NbFamilles
                TotalFamille += wsExcel.Cells(i + 2, j + 3).Value
            Next i
            wsExcel.Cells(NbFamilles + 3, j + 3).Value = TotalFamille

        Next j
        AlphaColTri = AlphaCol(NbCat + 4)
        With wsExcel.Cells("D" & NbFamilles + 3 & ":" & AlphaColTri & NbFamilles + 3).Style
            .HorizontalAlignment = ExcelHorizontalAlignment.Center
            .Font.Bold = True
        End With

        '------------------Mise en forme onglet AIDA------------------------------------------

        i = 2
        AlphaColTri = AlphaCol(NbCat + 3)
        While i < NbFamilles + 3         'colorie les lignes une sur deux
            With wsExcel.Cells("A" & i & ":" & AlphaColTri & i).Style
                .Fill.PatternType = ExcelFillStyle.Solid
                .Fill.BackgroundColor.Indexed = 26
            End With
            i += 2
        End While

        '---formatage des colonnes en kilo ou autre-------------------------------------------------
        For i = 1 To NbCat
            AlphaColTri = AlphaCol(i + 3)
            If UnAIDA(i) = "[kgC]" Or UnAIDA(i) = "[kgM]" Then
                FormatCol = "###0.0#;;#"
                wsExcel.Columns(i + 3).Style.Numberformat.Format = FormatCol
            Else
                FormatCol = "####;;#"
                wsExcel.Columns(i + 3).Style.Numberformat.Format = FormatCol
            End If
        Next

        AlphaColTri = AlphaCol(NbCat + 4)       'colonne Totaux
        FormatCol = "#0.0#;;#"
        wsExcel.Columns(NbCat + 4).Style.Numberformat.Format = FormatCol

        '--------------Codes barres ------------------------------------------------
        Decal = 3
        For j = 1 To NbCat
            jdec = j + Decal
            Contenu = CatAIDA(j)

            Call CodeBarreBMP(j, Contenu)

            wsExcel = wbExcel.Worksheets("AIDA")
            ' wsExcel.Cells(1, jdec).Select()

            Dim dT, dL, dW, dH As Single
            dT = 0                                   'coordonn�es du haut de l'image
            dL = 214 + (j - 1) * 91         ' coordonn�es du cot� gauche de l'image
            dW = 91                              ' largeur de l'image
            dH = 213                        ' hauteur de l'image

            Dim fileName = "Image" & j & ".png"
            Dim picture = wsExcel.Drawings.AddPicture(fileName, Path.Combine(CheminBureau, fileName))
            picture.SetSize(dW, dH)
            picture.SetPosition(dT, dL)
        Next

    End Sub

    Sub TestSomme2(nbden As Integer)
        '  ----- teste les quantit�s attribu�es------------------
        Dim j As Integer
        Dim AlphaColTri As String
        Dim Total As Single
        Dim i As Integer
        Dim Intitule() As String
        Dim Complet As String
        Dim Separ As Char = "("
        Dim Quant As Single
        Dim Brut As String
        Dim NbOcc As Integer

        For j = 1 To nbden
            Total = 0

            ' ------ d�codage de l'ent�te de colonne pour retrouver la quantit� d�clar�e
            Complet = wsExcel.Cells(1, j + Decal).Value             ' reprend l'ent�te de colonne
            NbOcc = CalculateOccurenceNumber(Complet, Separ)        ' compte le nombre de s�parateur dans l'ent�te
            Intitule = Complet.Split(Separ)                         ' �clate l'ent�te avec le s�parateur
            Brut = Intitule(NbOcc)                                  ' prend le string apr�s le dernier s�parateur
            Quant = Single.Parse(Brut.Remove(Brut.Length - 1, 1))   ' conversion du string en quantit� apr�s avoir enlev� la parenth�se

            ' ------ calcule le total des quantit�s attribu�es ------------
            For i = 1 To NbFamilles
                Total += wsExcel.Cells(i + 1, j + Decal).Value
            Next i

            ' ----- teste la valeur --------------------------------
            AlphaColTri = AlphaCol(j + Decal)
            Select Case Total
                Case 0      ' valeur nulle, rien de document�
                    TexteMsg = "La somme de la colonne " & AlphaColTri & " est nulle"
                    Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
                Case Quant      ' valeur �gale � la quantit� d�clar�e, RAS
                Case Else       ' valeur diff�rente = alerte
                    TexteMsg = "Col " & AlphaColTri & " " & Intitule(0) & ":  Somme= " & Total & " diff�rente de la quantit� d�clar�e" & Quant
                    Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RESULTATS")
                    NbErreur += 1
            End Select
        Next j

    End Sub
    Function CalculateOccurenceNumber(strString As String, strCharacter As String) As Integer

        Dim intPosition As Integer
        intPosition = 1
        While intPosition <= Len(strString) And strCharacter <> "" And InStr(intPosition, strString, strCharacter) <> 0
            intPosition = InStr(intPosition, strString, strCharacter) + 1
            CalculateOccurenceNumber += 1
        End While
        Return CalculateOccurenceNumber
    End Function


    Sub TestSomme(nbden As Integer)
        '  ----- test si une colonne est vide ------------------
        Dim j As Integer
        Dim AlphaColTri As String
        Dim Total As Single
        Dim i As Integer

        For j = 1 To nbden
            Total = 0
            For i = 1 To NbFamilles
                Total += wsExcel.Cells(i + 1, j + Decal).Value
            Next i
            AlphaColTri = AlphaCol(j + Decal)
            If Total = 0 Then
                TexteMsg = "La somme de la colonne " & AlphaColTri & " est nulle"
                Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RESULTATS")
                NbErreur += 1
            End If
        Next j

    End Sub
    Sub ReportCumul(nbdenrees As Integer, Decal As Integer, i As Integer, CodePrixDenree() As String, Categorie() As String,
    PoidsCat() As Single, Poids() As Single, UnAIDA() As String, Equiv() As Single)

        ' relecture des resultats, calcul des cumuls, denr�e par denr�e, � chaque cat�gorie en fonction de l'unit�
        ' si l'unit� n'est pas reconnue, sortie sous Erreur = True

        Dim j As Integer
        Dim k As Integer
        Dim CatErreur As String = ""
        Dim UnitErreur As String = ""
        Dim Erreur As Boolean

        wsExcel = wbExcel.Worksheets("RESULTATS")
        'Console.WriteLine("nbdenrees" & nbdenrees & "  nbcat " & NbCat)
        'Dim StrOption As String = Console.ReadLine()
        For j = 1 To nbdenrees
            Erreur = False
            For k = 1 To NbCat
                Select Case UCase(UnAIDA(k))
                    Case "[BOI]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).Value * Equiv(j)
                    Case "[KGC]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + Poids(j) * wsExcel.Cells(i + 1, j + Decal).Value
                    Case "[KGM]"
                        If Categorie(k) = CodePrixDenree(j) And wsExcel.Cells(NbFamilles + 2, j + Decal).Value <> 0 Then
                            PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).Value * Poids(j) / wsExcel.Cells(NbFamilles + 2, j + Decal).Value
                        End If
                    Case "[UN]"
                        If Categorie(k) = CodePrixDenree(j) Then PoidsCat(k) = PoidsCat(k) + wsExcel.Cells(i + 1, j + Decal).Value
                    Case Else
                        Erreur = True
                        CatErreur = Categorie(k)
                        UnitErreur = UnAIDA(k)
                End Select
            Next k
            If Erreur Then
                TexteMsg = "(ReportCumul) Pour les Codes Prix " & CatErreur & ", l'unit� " & UnitErreur & " n'est pas reconnue"
                Call Reporting("RESULTATS", "ALERTE", TexteMsg, "RAPPORT")
                NbErreur += 1
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

    Sub VerifPoids(j As Integer, nbprix As Integer, codeprixDenree As String, codeprix() As String, unitAida() As String, PoidsVerif As Single, TestPrepa As String, Onglet As String)
        Dim i As Integer
        '---- v�rification poids total non nul si unit� kgM -----------------------------------------------------
        For i = 1 To nbprix
            If codeprixDenree = codeprix(i) Then
                If unitAida(i) = "[kgM]" And PoidsVerif = 0 Then
                    TexteMsg = "Ligne " & j + 1 & " " & TestPrepa & ": unit� " & unitAida(i) & " et Poids total nul, pas de valeur dans la col. AIDA  " & codeprixDenree
                    Call Reporting(Onglet, "ALERTE", TexteMsg, Onglet)
                    NbErreur += 1
                End If
            End If
        Next i
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
        Dim IntHauteur As Integer = 350          ' Hauteur de l'image
        Dim IntLargeur As Integer = 900          ' Largeur de l'image
        Dim IntHautModule As Integer = 250       ' hauteur des barres
        Dim IntLargModule As Integer = 6       ' largeur des barres
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
        If IntLargeur < lngLargeurCodeBarres Then IntLargeur = CInt(lngLargeurCodeBarres * 1.05)
        IntMargeGauche = CInt((IntLargeur - lngLargeurCodeBarres) / 2)

        'TexteMsg = "Cha�ne " & Len(strChaine) & " Code Barre " & Len(strCodeBarres) & " intlargeur " & IntLargeur & "  marge " & IntMargeGauche
        'Call Reporting("Code", "ALERTE", TexteMsg, "AIDA")

        ' *******************************************************************
        ' cr�ation du graphique
        '********************************************************************
        Dim newBitmap As New SKBitmap(IntHauteur, IntLargeur, SKColorType.Rgba8888, SKAlphaType.Premul) 'cr�ons un BitMap vertical (inversion de hauteur et largeur)
        Dim g As New SKCanvas(newBitmap) 'cr�ons un Graphics et y mettre le BitMap
        'rotation de 270 de tous les �l�ments � dessiner
        g.Translate(IntHauteur / 2, IntLargeur / 2)
        g.RotateDegrees(-90)
        g.Translate(-IntLargeur / 2, -IntHauteur / 2)

        Dim blackPen As New SKPaint() 'cr�er un stylet noir d'�paisseur
        With blackPen
            .Color = New SKColor(0, 0, 0, 255)
            .StrokeWidth = 5
        End With
        Dim WhitePen As New SKPaint()
        With WhitePen
            .Color = New SKColor(255, 255, 255, 255)
            .StrokeWidth = 5
        End With

        g.Clear(New SKColor(255, 255, 255, 255))

        X1 = IntMargeGauche
        X2 = X1
        Y1 = 330
        Y2 = Y1 - IntHautModule

        '------Cr�ation de la zone de texte ------------------------------------------

        ' Dim drawFont As New System.Drawing.Font("Arial", 35)
        ' Dim drawBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Black)

        Dim drawFont As New SKFont(SKTypeface.FromFamilyName("Arial"), 50)
        Dim textPaint As New SKPaint(drawFont)
        textPaint.Color = New SKColor(0, 0, 0, 255)
        g.DrawText(Contenu, 250, 50, textPaint)

        '-----Tra�age du code-barres----------------------------------------------
        For i = 1 To Len(strCodeBarres)
            strTypeModule = Mid(strCodeBarres, i, 1)                                'Type de module, barre ou espace, � tracer
            Select Case strTypeModule
                Case "1"
                    g.DrawLine(X1, Y1, X2, Y2, blackPen)
                    X1 += IntLargModule
                    X2 = X1

                Case "0"
                    g.DrawLine(X1, Y1, X2, Y2, WhitePen)
                    X1 += IntLargModule
                    X2 = X1
            End Select

        Next i

        Using data = newBitmap.Encode(SKEncodedImageFormat.Png, 100)
            Using writer = File.OpenWrite(Path.Combine(CheminBureau, "Image" & j & ".png"))
                data.SaveTo(writer)
            End Using
        End Using

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
            Call Reporting("AIDA", "ALERTE", "Cha�ne de caract�res Code-Barres nulle ", "AIDA")
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

        MotifCodeBarres128 = ""
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

        Return MotifCodeBarres128
        'GestionErreurs:

        'Transmet l'erreur � la proc�dure appelante
        'Err.Raise(Err.Number, "MotifCodeBarres128")

    End Function
End Module

