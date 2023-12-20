Sub ImportArray()
    Dim DataBeatle As Variant
    Dim DataDemand As Variant
    Dim DataInfo As Variant
    Dim wsBeatle As Worksheet
    Dim wsBeatleDone As Worksheet
    Dim wsDemand As Worksheet
    Dim wsInfo As Worksheet
    Dim i As Long, j As Long, K As Long, L As Long, M As Long, n As Long, o As Long
    Dim matchFound As Boolean
    Dim Line As Long
    Dim iOrigRange As Integer
    Dim iDestRange As Integer
    Dim DestRange As Long
    Dim RowUpdate As Long
    Dim next_id As Boolean
    Dim TempDataDemand() As Variant
    Dim Col As Long
    Dim Row As Long

    ' Ajustar as referências das planilhas para corresponder aos nomes no seu workbook
    Set wsBeatleDone = ThisWorkbook.Worksheets("BEATLE_ASA_DONE")
    Set wsBeatle = ThisWorkbook.Worksheets("BEATLE_ASA")
    Set wsDemand = ThisWorkbook.Worksheets("SP Demand")
    Set wsInfo = ThisWorkbook.Worksheets("Macro_info")

    ' Índices da planilha de origem
    OrigID = 1
    OrigProjectCode = 2
    OrigRequest = 3
    OrigArea = 4
    OrigTemplate = 5
    OrigActivity = 6
    OrigSubactivity = 7
    OrigTotalQty = 8
    OrigAnalyticalDemand = 9
    OrigClientNotes = 10
    OrigPriority = 11
    OrigRushReason = 12
    OrigRequestStatus = 13
    OrigSamplesOrigin = 14
    OrigSamplesReceiptDate = 15
    OrigReceivedOn = 16
    OrigTIME_MIN = 17
    OrigDemandedDate = 18
    OrigPlannedDate = 19
    OrigFIRST_AGREED_DATE = 20
    OrigAGREED_DATE = 21
    OrigCLOSED_ON = 22
    OrigPITCode = 23
    OrigProjectName = 24
    OrigLEADER = 25
    OrigLYRA_REQUEST = 26
    OrigLYRA_CYCLE = 27
    OrigLYRA_PROJECT = 28
    OrigLYRA_LEADER = 29

    ' Índices da planilha destino
    DestID = 1
    DestProjectCode = 2
    DestRequest = 3
    DestArea = 4
    DestTemplate = 5
    DestActivity = 6
    DestSubactivity = 7
    DestTotalQty = 8
    DestAnalyticalDemand = 9
    DestClientNotes = 10
    DestPriority = 11
    DestRushReason = 12
    DestRequestStatus = 13
    DestSamplesOrigin = 14
    DestSamplesReceiptDate = 15
    DestDemandObservations = 16
    DestReceivedOn = 17
    DestProductionStatus = 18
    DestPlanningNotes = 19
    DestProductionQty = 20
    DestTimemin = 21
    DestPlannedProductionDate = 22
    DestRealProductionDate = 23
    DestDeliveryDate = 24
    DestMissClassification = 25
    DestSPNotes = 26
    DestDemandedDate = 27
    DestPlannedDate = 28
    DestFirstAgreedDate = 29
    DestAgreedDate = 30
    DestClosedOn = 31
    DestPITCode = 32
    DestProjectName = 33
    DestLeader = 34
    DestLyraRequest = 35
    DestUpdateObs = 36
    DestLyraProject = 37
    DestLyraLeader = 38
    DestAcordodePrazo = 39
    DestEntreganolab = 40

    Query = 1
    Do
        If Query = 1 Then
            ' Lendo dados da planilha BEATLE_ASA
            DataBeatle = wsBeatle.Range("A1", wsBeatle.Cells(wsBeatle.Rows.Count, "A").End(xlUp)).Value
        Else
            ' Lendo dados da planilha BEATLE_ASA_DONE
            DataBeatle = wsBeatleDone.Range("A1", wsBeatleDone.Cells(wsBeatleDone.Rows.Count, "A").End(xlUp)).Value
        End If



        ' Lendo dados da planilha SP Demand
        With wsDemand
            DataDemand = .Range("A1", .Cells(.Rows.Count, "A").End(xlUp)).Value
        End With

        ' Lendo dados da planilha Macro_info
        With wsInfo
            DataInfo = .Range("A1", .Cells(.Rows.Count, "A").End(xlUp)).Value
        End With


        ' Supondo que DataDemand e DataBeatle já tenham sido preenchidos com dados
            

                ' Comparando com cada item na primeira coluna do DataDemand
            For i = LBound(DataBeatle, 1) To UBound(DataBeatle, 1)
                itsNew = True ' Assume que é novo por padrão

                For j = LBound(DataDemand, 1) To UBound(DataDemand, 1)
                    If DataBeatle(i, 1) = DataDemand(j, 1) Then
                        itsNew = False
                        RowUpdate = j
                        Exit For
                    End If
                Next j

                If itsNew Then
                    lastRow = UBound(DataDemand, 1)
                    ReDim Preserve DataDemand(1 To lastRow + 1, 1 To UBound(DataDemand, 2))
                    For Col = LBound(DataDemand, 2) To UBound(DataDemand, 2)
                        DataDemand(lastRow + 1, Col) = "" ' Ou algum valor padrão
                    Next Col
                    RowUpdate = lastRow + 1
                End If
                    
        
                    If itsNew = True Then

                        DestRange = LBound(DataDemand, 1) + 1 'posição para inserir dado na última linha do Array
                        itsNew = True

                    'Não é uma demanda nova
                    Else

                    
                        itsNew = False
                    End If

                        Line = 2
                        ID = Empty
                        Template = Empty
                        Activity = Empty
                        Subactivity = Empty
                        ReceiptDate = Empty
                        ReceivedOn = Empty
                        base_date = Empty
                        PlannedProdDate = Empty
                        Update_Status = False
                        Update_Date = False
                        Update_Obs = ""


                        Do 'Loop

                            If itsNew = True Then
                                goin = True
                                Update_Status = True
                                Update_Date = True
                            Else
                                If DataInfo(Line, 3) = "T" Then
                                    goin = True
                                Else
                                    goin = False
                                End If
                            End If

                            If goin = True Then

                                If Line >= 2 And Line <= 16 Then
                                    iOrigRange = Line - 1
                                ElseIf Line = 18 Then
                                    iOrigRange = Line - 2
                                ElseIf Line = 22 Then
                                    iOrigRange = Line - 5
                                ElseIf Line >= 28 And Line <= 36 Then
                                    iOrigRange = Line - 10
                                End If

                                If Line >= 2 And Line <= 39 Then
                                    iDestRange = Line - 1
                                End If

                                If DataInfo(Line, 2) <> "" Then

                                    Not_Update_Cell = False
                                
                                    If DataInfo(Line, 3) = "ID" Then
                                        ID = DataBeatle(i, OrigID)
                                        Not_Update_Cell = True
                                    ElseIf DataInfo(Line, 3) = "Template" Then
                                        Template = DataBeatle(i, OrigTemplate)
                                        Not_Update_Cell = True
                                    ElseIf DataInfo(Line, 3) = "Activity" Then
                                        Activity = DataBeatle(i, OrigActivity)
                                        Not_Update_Cell = True
                                    ElseIf DataInfo(Line, 3) = "Subactivity" Then
                                        Subactivity = DataBeatle(i, OrigSubactivity)
                                        Not_Update_Cell = True
                                    ElseIf DataInfo(Line, 3) = "ReceiptDate" Then
                                        ReceiptDate = DataBeatle(i, OrigSamplesReceiptDate)
                                    ElseIf DataInfo(Line, 3) = "ReceivedOn" Then
                                        ReceivedOn = DataBeatle(i, OrigReceivedOn)
                                    End If

                                    If itsNew = False Then 'não é uma demanda nova, checar e fazer possíveis updates
                                        If DataInfo(Line, 5) = "T" Then 'Atualizar o status
                                            If DataDemand(RowUpdate, iDestRange) <> DataBeatle(i, iOrigRange) Then
                                                If DataInfo(Line, 2) <> "ReceivedOn" Then
                                                    Update_Status = True
                                                    Update_Obs = Update_Obs & " " & DataInfo(Line, 1)
                                                End If
                                            End If
                                        End If

                                        If DataInfo(Line, 6) = "T" Then 'Atualiza Datas
                                            If DataInfo(Line, 3) = "ReceivedOn" Then
                                                If DataDemand(RowUpdate, iDestRange) <> DataBeatle(i, iOrigRange) And ReceiptDate <> DataBeatle(i, iOrigRange) Then
                                                    Update_Date = True
                                                    Update_Status = True
                                                    Update_Obs = Update_Obs & " " & DataInfo(Line, 1)
                                                End If
                                            Else
                                                If DataDemand(RowUpdate, iDestRange) <> DataBeatle(i, iOrigRange) Then
                                                    Update_Date = True
                                                End If
                                            End If

                                        End If

                                        If Not_Update_Cell = False Then
                                            DataDemand(RowUpdate, iDestRange) = DataBeatle(i, iOrigRange)
                                        End If

                                    Else
                                        DataDemand(RowUpdate, iDestRange) = DataBeatle(i, iOrigRange)
                                    End If
                                
                                ElseIf DataInfo(Line, 3) = "Status" Then

                                    If Update_Status = True Then
                                        If DataDemand(RowUpdate, iDestRange) <> "Finished" And DataDemand(RowUpdate, iDestRange) <> "Confirmed" And DataDemand(RowUpdate, iDestRange) <> "Planned_CR" And DataDemand(RowUpdate, iDestRange) <> "Cancelled" Then
                                            DataDemand(RowUpdate, iDestRange) = DataInfo(Line, 7)
                                        Else
                                            Update_Date = False
                                        End If
                                    End If

                                ElseIf DataInfo(Line, 3) = "CalcPlannedProdDate" Then

                                    If Update_Date = True Then

                                        If ReceivedOn = Empty Then
                                            On Error Resume Next
                                            base_date = ReceiptDate
                                        Else
                                            On Error Resume Next
                                            base_date = ReceivedOn
                                        End If

                                        If Activity = "Eletronic Receipt" Then
                                                num_days = 1
                                        ElseIf Activity = "Cigarette Preparation" Then
                                            If Subactivity = "Material Separation" Then
                                                num_days = 1
                                            Else
                                                num_days = 2
                                            End If
                                        ElseIf Activity = "Tobacco Preparation" Then
                                            If Template = "BR_PROTOTYPES" Then
                                                num_days = 2
                                            Else
                                                num_days = 3
                                            End If
                                        End If

                                        If Template = "BR_PMD_TOBACCO_CONTROL" Then
                                            PlannedProdDate = base_date
                                        Else
                                            PlannedProdDate = Application.WorkDay(base_date, num_days, Range("Holidays"))
                                        End If

                                        DataDemand(RowUpdate, iDestRange) = PlannedProdDate

                                    End If
                                            
                                ElseIf DataInfo(Line, 3) = "CalcDeliveryDate" Then

                                    If Update_Date = True Then
                                    
                                        update_delivery = True
                                        
                                        If Activity = "Eletronic Receipt" Then
                                            update_delivery = False
                                            position = InStr(1, ID, ".")

                                            Filter_text = Left(ID, position - 1) & ".2.?"
                                            ' Inicializa a variável 'found' como False
                                            next_id = False

                                            ' Loop para percorrer apenas a primeira coluna do array
                                            For L = LBound(DataBeatle, 1) To UBound(DataBeatle, 1) ' Itera sobre a primeira dimensão (linhas)
                                                If DataBeatle(i, 1) = Filter_text Then ' Compara o valor na primeira coluna
                                                    next_id = True
                                                    Exit For
                                                End If
                                            Next L

                                            If next_id = False Then
                                                update_delivery = True
                                            End If

                                            Filter_text = Left(ID, position - 1) & ".3.?"
                                            
                                            For M = LBound(DataBeatle, 1) To UBound(DataBeatle, 1) ' Itera sobre a primeira dimensão (linhas)
                                                If DataBeatle(i, 1) = Filter_text Then ' Compara o valor na primeira coluna
                                                    next_id = True
                                                    Exit For
                                                End If
                                            Next M

                                            If next_id = False Then
                                                update_delivery = True
                                            End If

                                            If update_delivery = True Then
                                            If Template = "BR_PMD_TOBACCO_CONTROL" Then
                                                DataDemand(RowUpdate, iDestRange) = PlannedProdDate
                                            Else
                                                DataDemand(RowUpdate, iDestRange) = Application.WorksheetFunction.WorkDay(PlannedProdDate, 1, Range("Holidays"))
                                            End If
                                        End If

                                    ElseIf DataInfo(Line, 1) = "Update Obs" Then
                                    If Update_Obs <> "" Then
                                        DataDemand(RowUpdate, iDestRange) = "Infos:" & Update_Obs
                                    End If
                                End If
                                
                            End If
                                End If

                                Line = Line + 1


                                ' Comparando com cada item na primeira coluna do DataDemand
                                For o = LBound(DataDemand, 1) To UBound(DataDemand, 1)
                                    If DataBeatle(i, 1) = DataDemand(o, 1) Then
                                        itsNew = False
                                    Else
                                        itsNew = True
                                        Exit For
                                    End If
                                Next o
                            On Error Resume Next
                            
                        Loop While DataInfo(Line, 1) <> "Lyra Leader"

            Next i

        Query = Query + 1
    Loop While Query < 3

    Sheets("BEATLE_ASA").Visible = False
    Sheets("BEATLE_ASA_DONE").Visible = False
    Sheets("Macro_info").Visible = False
    ' 1. Limpar os dados existentes na planilha "SP Demand"
    With wsDemand
        .Range("A1", .Cells(.Rows.Count, "A").End(xlUp)).ClearContents
    End With

    ' 2. Escrever os dados de DataDemand de volta para a planilha "SP Demand"
    With wsDemand
        Dim lastRow As Long
        lastRow = UBound(DataDemand, 1)
        Dim lastCol As Long
        lastCol = UBound(DataDemand, 2)

        ' Define o range na planilha que corresponderá ao tamanho do array DataDemand
        Dim DestRange As Range
        Set DestRange = .Range("A1").Resize(lastRow, lastCol)

        ' Atribui os dados do array para a planilha
        DestRange.Value = DataDemand
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox ("Planilha Atualizada")

    End Sub


        


