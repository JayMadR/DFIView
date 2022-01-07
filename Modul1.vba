Sub getCurrentValues()
    'Wosksheet inputs
    Dim wksMaster As Worksheet
    Set wksMaster = Worksheets("myLT")


    'Quantity APIs
    Const iStart As Integer = 0
    Const iQty As Integer = 5


    'inactive Indizes (0 = active | else -> not active)
    Dim iNotAct(iStart To iQty) As Integer
    iNotAct(0) = 0
    iNotAct(1) = 0
    iNotAct(2) = 0
    iNotAct(3) = 0
    iNotAct(4) = 0
    iNotAct(5) = 0
    
    'col/row change (0 = normal | else -> change)
    Dim irowChange(iStart To iQty) As Integer
    irowChange(0) = 0
    irowChange(1) = 0
    irowChange(2) = 0
    irowChange(3) = 0
    irowChange(4) = 0
    irowChange(5) = 1
    
    'with Index (0 = active | else -> not active)
    Dim iIdx(iStart To iQty) As Integer
    iIdx(0) = 0
    iIdx(1) = 0
    iIdx(2) = 1
    iIdx(3) = 0
    iIdx(4) = 1
    iIdx(5) = 0
    
    'structString
    Dim iStructStr(iStart To iQty) As String
    iStructStr(0) = "data"
    iStructStr(1) = "data"
    iStructStr(2) = "data"
    iStructStr(3) = "data"
    'iStructStr(4) = ""
    iStructStr(5) = "data"
    
    
    'API - Links
    Dim APILink(iStart To iQty) As String
    APILink(0) = "https://ocean.defichain.com/v0/mainnet/poolpairs?size=1000"
    APILink(1) = "https://ocean.defichain.com/v0/mainnet/prices?size=1000"
    APILink(2) = "https://ocean.defichain.com/v0/mainnet/stats"
    APILink(3) = "https://ocean.defichain.com/v0/mainnet/address/XXXXX/tokens"
    APILink(4) = "https://api.binance.com/api/v3/ticker/price?symbol=BTCEUR"
    APILink(5) = "https://ocean.defichain.com/v0/mainnet/address/XXXXX/vaults"
    
    'Insert Address into API-Link
    APILink(3) = Replace(APILink(3), "XXXXX", wksMaster.Cells(4, 5).Value)
    APILink(5) = Replace(APILink(5), "XXXXX", wksMaster.Cells(4, 5).Value)
    
    
    'Worksheets
    Dim wks As Worksheet
    Dim sheets(iStart To iQty) As String
    sheets(0) = "PoolPairs"
    sheets(1) = "Prices"
    sheets(2) = "Stats"
    sheets(3) = "Address"
    sheets(4) = "BTCEur"
    sheets(5) = "Vaults"
    
    
    'Set Keys
    Dim keys(iStart To iQty) As Variant
    keys(0) = Array("idx", "id", "symbol", "displaySymbol", "name", "tokenA.symbol", "tokenA.displaySymbol", "tokenA.id", "tokenA.reserve", "tokenA.blockComission", "tokenB.symbol", "tokenB.displaySymbol", "tokenB.id", "tokenB.reserve", "tokenB.blockComission", "priceRatio.ab", "priceRatio.ba", "totalLiquidity.token", "totalLiquidity.usd", "apr.reward", "apr.total", "commission", "rewardPct", "status", "tradeEnabled", "ownerAddress", "creation.tx", "creation.height")
    keys(1) = Array("idx", "id", "sort", "price.currency", "price.token", "price.id", "price.key", "price.sort", "price.aggregated.amount", "price.aggregated.weightage", "price.aggregated.oracles.active", "price.aggregated.oracles.total", "price.block.hash", "price.block.height", "price.block.medianTime", "price.block.time")
    keys(2) = Array("idx", "count.blocks", "emission.masternode", "emission.dex", "emission.community", "emission.anchor", "emission.burned", "emission.total", "tvl.total")
    keys(3) = Array("idx", "id", "amount", "symbol", "symbolKey", "name", "isDAT", "isLPS", "displaySymbol")
    keys(4) = Array("idx", "symbol", "price")
    keys(5) = Array("idx", "vaultId", "loanScheme.id", "loanScheme.minColRatio", "loanScheme.interestRate", "ownerAddress", "state", "informativeRatio", "collateralRatio", "collateralValue", "loanValue", "interestValue", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    
    
    Dim n As Integer
    n = 12
    For i = 1 To 4
        keys(5)(n) = "collateralAmounts." & i & ".id"
        n = n + 1
        keys(5)(n) = "collateralAmounts." & i & ".symbol"
        n = n + 1
        keys(5)(n) = "collateralAmounts." & i & ".amount"
        n = n + 1
        keys(5)(n) = "collateralAmounts." & i & ".activePrice.isLive"
        n = n + 1
        keys(5)(n) = "collateralAmounts." & i & ".activePrice.active.amount"
        n = n + 1
        keys(5)(n) = "collateralAmounts." & i & ".activePrice.next.amount"
        n = n + 1
    Next
    
    For i = 1 To 16
        keys(5)(n) = "loanAmounts." & i & ".id"
        n = n + 1
        keys(5)(n) = "loanAmounts." & i & ".symbol"
        n = n + 1
        keys(5)(n) = "loanAmounts." & i & ".amount"
        n = n + 1
        keys(5)(n) = "interestAmounts." & i & ".amount"
        n = n + 1
        keys(5)(n) = "loanAmounts." & i & ".activePrice.isLive"
        n = n + 1
        keys(5)(n) = "loanAmounts." & i & ".activePrice.active.amount"
        n = n + 1
        keys(5)(n) = "loanAmounts." & i & ".activePrice.next.amount"
        n = n + 1
    Next
    
    


    'received Data (Json-Format)
    Dim AllData(iStart To iQty) As Object
    Dim data As Object
    
    
    'Loop -> get Data from APIs
    For i = iStart To iQty
        If iNotAct(i) = 0 Then
            'get Data
            Set AllData(i) = getJSON(APILink(i))
        End If
    Next
    
    
    'Rows/Cols
    Dim row As Integer
    Dim col As Integer
    Dim initRow As Integer
    Dim initCol As Integer
    initRow = 5
    initCol = 2
    Dim key As Variant
    Dim noEntry As Integer
    noEntry = 0
    
    
    'Loop -> ClearContents
    For i = iStart To iQty
        'Worksheet zuweisen
        Set wks = Worksheets(sheets(i))
                        
        'Delete
        wks.Range(wks.Cells(initRow, initCol), wks.Cells(1000, 1000)).ClearContents
    Next
    
    
    'Loop -> Output Data
    For i = iStart To iQty
        If iNotAct(i) = 0 Then
            'Set Worksheet
            Set wks = Worksheets(sheets(i))
            
            'row/col init
            row = initRow
            col = initCol
        
            'write API-link (as INFO)
            wks.Cells(row - 1, col).Value = "API"
            wks.Cells(row - 1, col + 1).Value = APILink(i)
        
            'write Labels (as INFO)
            For Each key In keys(i)
                wks.Cells(row, col).Value = key
                'col = col + 1
                If irowChange(i) = 0 Then
                    col = col + 1
                Else
                    row = row + 1
                End If
                '----------------
            Next key
            'row = row + 1
            If irowChange(i) = 0 Then
                row = row + 1
            Else
                col = col + 1
            End If
            '----------------
        
            'wks.Range(wks.Cells(row, initCol), wks.Cells(1000, col + 1)).ClearContents
            If irowChange(i) = 0 Then
                wks.Range(wks.Cells(row, initCol), wks.Cells(1000, col + 1)).ClearContents
            Else
                wks.Range(wks.Cells(initRow, initCol + 1), wks.Cells(1000, 1000)).ClearContents
            End If
            '----------------
        
            'copy Data
            Set data = AllData(i)
        
            Dim subkeys As Integer
            'write data
            For j = 1 To 100
                'col = initCol
                If irowChange(i) = 0 Then
                    col = initCol
                Else
                    row = initRow
                End If
                '----------------
                For Each key In keys(i)
                    On Error GoTo ErrorHandler
                    If key = keys(i)(0) Then
                        wks.Cells(row, col).Value = CStr(j)
                    Else
                        'BTC-EUR
                        If i = 4 Then
                            wks.Cells(row, col).Value = data(key)
                        Else
                            'Sub-Key
                            wks.Cells(row, col).Value = getDatum(data, iStructStr(i), key, j, iIdx(i))
                        End If
                    End If
                    'col = col + 1
                    If irowChange(i) = 0 Then
                        col = col + 1
                    Else
                        row = row + 1
                    End If
                    '----------------
                Next key
                'row = row + 1
                If irowChange(i) = 0 Then
                    row = row + 1
                Else
                    col = col + 1
                End If
                '----------------
                'no entries left
                If iIdx(i) = 1 Then
                    If row > initRow + 2 Then
                        'wks.Cells(row - 1, initCol).Value = ""
                        If irowChange(i) = 0 Then
                            wks.Cells(row - 1, initCol).Value = ""
                        Else
                            'what here?
                        End If
                    End If
                    Exit For
                End If
            Next
        End If
    Next
    
    Application.Calculate
    
    

'CInt -> String to int

ErrorHandler:
    'Datum in Data not found
    If Err.Number = 9 Then
        'Exit For
        iIdx(i) = 1
    End If
Resume Next


End Sub




Function getDatum(dataIn As Object, structStr As String, key As Variant, j As Variant, idx As Integer) As Variant
    Dim splitter As String
    splitter = "."
    Dim subkey(0 To 5) As Variant

    'with Idx
    If idx = 0 Then
        If InStr(key, splitter) > 0 Then
            subkeys = Len(key) - Len(Replace(key, splitter, ""))
            
            'get substrings
            For m = 0 To subkeys
                subkey(m) = split(key, splitter)(m)
                'if IsNumeric -> toInt
                If IsNumeric(subkey(m)) Then
                    subkey(m) = CInt(subkey(m))
                End If
            Next
            
            Select Case subkeys
                Case 1
                    getDatum = dataIn(structStr)(j)(subkey(0))(subkey(1))
                Case 2
                    getDatum = dataIn(structStr)(j)(subkey(0))(subkey(1))(subkey(2))
                Case 3
                    getDatum = dataIn(structStr)(j)(subkey(0))(subkey(1))(subkey(2))(subkey(3))
                Case 4
                    getDatum = dataIn(structStr)(j)(subkey(0))(subkey(1))(subkey(2))(subkey(3))(subkey(4))
                Case 5
                    getDatum = dataIn(structStr)(j)(subkey(0))(subkey(1))(subkey(2))(subkey(3))(subkey(4))(subkey(5))
                Case Else
                    'WHAT DO DO HERE?
            End Select
        Else
            getDatum = dataIn(structStr)(j)(key)
        End If
    Else
        If InStr(key, splitter) > 0 Then
            subkeys = Len(key) - Len(Replace(key, splitter, ""))
                        
            'get substrings
            For m = 0 To subkeys
                subkey(m) = split(key, splitter)(m)
            Next
            
            Select Case subkeys
                Case 1
                    getDatum = dataIn(structStr)(subkey(0))(subkey(1))
                Case 2
                    getDatum = dataIn(structStr)(subkey(0))(subkey(1))(subkey(2))
                Case 3
                    getDatum = dataIn(structStr)(subkey(0))(subkey(1))(subkey(2))(subkey(3))
                Case 4
                    getDatum = dataIn(structStr)(subkey(0))(subkey(1))(subkey(2))(subkey(3))(subkey(4))
                Case 5
                    getDatum = dataIn(structStr)(subkey(0))(subkey(1))(subkey(2))(subkey(3))(subkey(4))(subkey(5))
                Case Else
                    'WHAT DO DO HERE?
            End Select
        Else
            getDatum = dataIn(structStr)(key)
        End If
    End If

End Function




'Function returns the whole JSON Object of the provided API link
Function getJSON(link As Variant) As Object
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    Request.Open "GET", link
    On Error Resume Next
    Request.Send
    
    'MsgBox Request.ResponseText
    
    Set getJSON = JsonConverter.ParseJson(Request.ResponseText)
End Function
