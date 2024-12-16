Sub TurnOffDuringProcess(App As Application, IsEnd As Boolean)
    With App
        If IsEnd <> True Then
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
        Else
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
        End If
    End With
End Sub

Sub FaireUnExtract()
    Dim SQL As New SQL_InVBA
    SQL.Class_Initialize
    SQL.ConnectionDB DB_Mapping.Range("PATH2")
    SQL.DesignRequest "Date", "Sheet1"
    SQL.AddToRequest " WHERE Famille = 'DS'"
    SQL.ExecuteRequest
    SQL.CloseDB
    
    Call ArrayToRange(OutPut, "A1", SQL.matriceGlobale, True)
    
    Dim matriceDates() As Variant
    Call ArrayToTransposedArray(SQL.matriceGlobale, matriceDates)
    Dim countdate As Integer
    countdate = GetCountDatesHowAnotherdate(matriceDates, Date)
    
    Call ArrayToRange(OutPut, "D1", matriceDates, False)
    
End Sub

Sub ArrayToRange(ws As Worksheet, location As String, matrice As Variant, transpose As Boolean) 'Easy : placer une matrice dans un worksheet
    Dim Lignes As Integer
    Dim Colonnes As Integer
    
    Lignes = IIf(UBound(matrice, 1) = 0, 1, UBound(matrice, 1)) - 1
    Colonnes = IIf(UBound(matrice, 2) = 0, 1, UBound(matrice, 2)) - 1
    If transpose = False Then
        ws.Range(ws.Range(location), ws.Range(location).Offset(Lignes, Colonnes)).ClearContents
        ws.Range(ws.Range(location), ws.Range(location).Offset(Lignes, Colonnes)) = matrice
    ElseIf transpose = True Then
        ws.Range(ws.Range(location), ws.Range(location).Offset(Colonnes, Lignes)).ClearContents
        ws.Range(ws.Range(location), ws.Range(location).Offset(Colonnes, Lignes)) = WorksheetFunction.transpose(matrice)
    End If
    
End Sub

Sub ArrayToTransposedArray(arrayStart As Variant, arrayEnd As Variant)
    Dim ligne As Integer
    Dim Colonne As Integer
    
    arrayEnd = WorksheetFunction.transpose(arrayStart)
End Sub

Function GetCountDatesHowAnotherdate(matrice As Variant, ComparaisonDate As Date) 'avoir un decompte des dates supérieures (>) à une autre date
    
    GetCountDatesHowAnotherdate = GetCountDatesHowAnotherdateWithSpecifiedColumn(matrice, ComparaisonDate, 1)
    
End Function


Function GetCountDatesHowAnotherdateWithSpecifiedColumn(matrice As Variant, ComparaisonDate As Date, Colonne As Integer) ' permet d'analyser les lignes qui représentent des dates et savoir lesquelles sont > à une certaine date
    Dim count As Integer
    Dim ligne As Integer
    Dim iter As Integer
    
    count = 0
    ligne = UBound(matrice, 1)
    
    For iter = 1 To ligne
        If CDate(matrice(iter, Colonne)) > CDate(ComparaisonDate) Then
            count = count + 1
        End If
    Next iter
    GetCountDatesHowAnotherdate = count
    
End Function

Function TurnUSDateIntoEUDate(texte As String) 'permet de traduire ce qui est écrit dans l'extraction ICE
    TurnUSDateIntoEUDate = Right(Left(texte, 5), 2) & "/" & Right(texte, 2) & "/" & Right(texte, 4)
End Function

Function QuantiteParFixingAndLeverage(quantite1 As Double, quantite2 As Double)
    Dim matriceretour() As Double
    ReDim matriceretour(1 To 3) As Double
    If quantite1 > quantite2 Then
        matriceretour(1) = quantite2
        matriceretour(2) = quantite1
        matriceretour(3) = matriceretour(2) / matriceretour(1)
    ElseIf quantite1 < quantite2 Then
        matriceretour(1) = quantite1
        matriceretour(2) = quantite2
        matriceretour(3) = matriceretour(2) / matriceretour(1)
    ElseIf quantite1 = quantite2 Then
        matriceretour(1) = quantite2
        matriceretour(2) = quantite1
        matriceretour(3) = 1
    End If
    QuantiteParFixingAndLeverage = matriceretour
End Function


Option Explicit
Option Base 1

Public connection As Object
Public request As String
Public recordset As Object
Public matriceGlobale As Variant
Public NBColonnes As Integer
Public NBLignes As Integer

Sub Class_Initialize()
    Debug.Print "Object initialisé"
    Set connection = CreateObject("ADODB.connection")
    Set recordset = CreateObject("ADODB.Recordset")
End Sub

Sub ConnectionDB(Path As String)
    connection.Provider = "Microsoft.ACE.OLEDB.16.0"
    connection.ConnectionString = "Data Source=" & Path & ";" & "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
    connection.Open
    Debug.Print "Connection à la base dont le chemin d'accès est : " & Path
End Sub

Sub DesignRequest(TexteDuSelect As String, NomFeuille As String)
    request = "SELECT " & TexteDuSelect & " FROM [" & NomFeuille & "$]"
    Debug.Print ("La request SQL est : " & request)
End Sub

Sub AddToRequest(texte As String)
    request = request & texte
    Debug.Print ("Ajoute de : " & texte)
    Debug.Print ("La requte totale est : " & request)
End Sub

Sub ExecuteRequest()
    request = request & ";"
    recordset.Open request, connection
    Debug.Print "Requete realisée"
    If Not recordset.EOF Then
        ' Extraction des données du Recordset dans une matrice
        
        matriceGlobale = recordset.GetRows()  ' Attention : les données sont transposées
        NBLignes = UBound(matriceGlobale, 1)
        NBColonnes = UBound(matriceGlobale, 2)
    End If
End Sub

Sub PrintOutput(Position As String)
    OutPut.Cells.Clear
    OutPut.Range(Position).CopyFromRecordset OutputClass
End Sub

Sub MettreDansMatrice(matrice As Variant)
    matrice = matriceGlobale
End Sub

Sub CloseDB()
    connection.Close
End Sub
