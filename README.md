# R2
Attribute VB_Name = "Module1"

'Variables para procesar
Dim hReport, hRedna As Worksheet
Dim LRRedna, LRReport As Long
Dim MRedna, MReport As Variant
Dim Today As String

'INFO DE T&T DATA 
Dim wbData As Workbook
Dim hTypes, hSekom, hTimes As Worksheet
Dim DTypes, DSekom, DAccounts, DHeaders As Object
Dim MTypes, MSekom, MTimes As Variant

Sub Test()

  'Información de este workbook
  Set hReport = ThisWorkbook.Sheets("Report")
  Set hRedna = ThisWorkbook.Sheets("Redna")

  'Información de T&T Data
  Set wbData = Workbooks.Open(ThisWorkbook.Path & "\T&T Data.xlsx")
  Set hTypes = wbData.Sheets("Type")
  Set hSekom = wbData.Sheets("Sekom")
  Set hTimes = wbData.Sheets("Times")

  'Obteniendo el indice de las ultimas filas
  LRTypes = hTypes.Cells(hTypes.Rows.Count, 1).End(xlUp).Row
  LRSekom = hSekom.Cells(hSekom.Rows.Count, 1).End(xlUp).Row
  LRTimes = hTimes.Cells(hTimes.Rows.Count, 2).End(xlUp).Row    
  LRRedna = hRedna.Cells(hRedna.Rows.Count, 1).End(xlUp).Row  
  LRReport = hReport.Cells(hReport.Rows.Count, 1).End(xlUp).Row

  'Inicializando matrices
  MTypes = hTypes.Range("A2:B" & LRTypes).Value
  MSekom = hSekom.Range("D2:E" & LRSekom).Value
  MTimes = hTimes.Range("B1:L" & LRTimes).Value

  'Creación de dictionarios
  Set DTypes = CreateObject("scripting.dictionary")
  Set DSekom = CreateObject("scripting.dictionary")
  Set DAccounts = CreateObject("scripting.dictionary")
  Set DHeaders = CreateObject("scripting.dictionary")
  Call SetDictionary(MTypes, DTypes)
  Call SetDictionary(MSekom, DSekom)
  Call SetDictionary2(MTimes, DAccounts)
  Call SetDictionary3(MTimes, DHeaders)
  wbData.Close SaveChanges:=False

  ' ==============================================================
  '              PASAR VALORES DE REDNA AL REPORT
  ' ==============================================================
  'OJOOOOOOOOOO CAMBIAR LA ASIGNACION POR NOMBRE DE COLUMNAS SERIA LO MEJOR
  hReport.Range("A" & LRReport + 1 & ":D" & (LRRedna + LRReport - 1)).Value = hRedna.Range("A2:D" & LRRedna).Value
  hReport.Range("F" & LRReport + 1 & ":J" & (LRRedna + LRReport - 1)).Value = hRedna.Range("E2:I" & LRRedna).Value
  hReport.Range("L" & LRReport + 1 & ":M" & (LRRedna + LRReport - 1)).Value = hRedna.Range("J2:K" & LRRedna).Value
  hReport.Range("O" & LRReport + 1 & ":AC" & (LRRedna + LRReport - 1)).Value = hRedna.Range("L2:Z" & LRRedna).Value
 

  ' ==============================================================
  '                       OBTENER "TYPE" y "ZONE"
  ' ==============================================================
  'Obtener Matriz de Reporte
  LRReport = hReport.Cells(hReport.Rows.Count, 1).End(xlUp).Row
  MReport = hReport.Range("A2:AC" & LRReport).Value

  'Buscar el indice de las columnas a usar
  ColStatus = Application.WorksheetFunction.Match("Status", hReport.Rows(1), 0)
  ColDestination = Application.WorksheetFunction.Match("Destination", hReport.Rows(1), 0)
  ColType = Application.WorksheetFunction.Match("TYPE", hReport.Rows(1), 0)
  ColZone = Application.WorksheetFunction.Match("ZONE", hReport.Rows(1), 0)

  'Ciclar por la matriz y buscar los valores
  For fila = 1 to UBound(MReport)
    status = MReport(fila, ColStatus)
    destination = MReport(fila, ColDestination)

    Type_ = GetDicValue(status, DTypes)         'Buscamos el type en el dictionary
    Zone = GetDicValue(destination, DSekom)     'Buscamos el color de sekom en el dictionary

    MReport(fila, ColType) = Type_
    MReport(fila, ColZone) = Zone
  Next fila

  'Actualizar el reporte
  hReport.Range("A2:AC" & LRReport).Value = MReport
  Call ApplyFormat(hReport)


  ' ==============================================================
  '                         OBTENER DELAYS
  ' ==============================================================
  'Ciclar por la matriz y buscar los valores
  Today = Format(Date, "mm/dd/yy")
  ColDelAppt = Application.WorksheetFunction.Match("Del Appt", hReport.Rows(1), 0)
  ColDueDate = Application.WorksheetFunction.Match("Due Date", hReport.Rows(1), 0)
  ColDTC = Application.WorksheetFunction.Match("DTC", hReport.Rows(1), 0)

  For fila = 1 to UBound(MReport)
    DelvAppt = Format( CDate(MReport(fila, ColDelAppt)), "mm/dd/yy")

    If MReport(fila, ColDelAppt) = "" Or DelvAppt < Today Then
      DTC = MReport(fila, ColType)

      'Analizamos si es un Type DTC
      If DTC = "X" Then 
        Type_ = "DTC"
      Else
        Type_ = MReport(fila, ColType) 
      End If

      'Analizamos si el Type se encontro con anterioridad o no
      If Type_ <> "Not Found" Then
        Account = "Weber"                                                           'OJOOOOO MANEJAR VARIACIONES
        Header = DHeaders.Item(Type_)

        'Debug.Print "TYPE: " & Type_ & "    HEADER: " & Header
        ColDate = Application.WorksheetFunction.Match(Header, hReport.Rows(1), 0)

        'Revisamos si la fecha de la columna(header) esta vacia o no
        If MReport(fila, ColDate) = "" And (Type_ = "O2" Or Type_ = "D2") Then
          Header = DHeaders.Item(Type_ & "_")
          ColDate = Application.WorksheetFunction.Match(Header, hReport.Rows(1), 0)
        End If

        'Obtenemos el string de Days y lo convertimos a entero
        key = Type_ & "|" & Header
        StringDays = DAccounts.Item(Account).Item(key)
        Days = CInt(Left(StringDays, Len(StringDays) - 1))                                            ' Quita la "D"
        DD = Days
        DateFrom = Format(Today, "mm/dd/yy")

        'Calculamos la fecha de inicio para evaluar
        'Restar días hasta que hayamos restado la cantidad de días laborables
        Do While Days > 0
            'DateFrom = DateFrom - 1  ' Restar un día
            'Restar un día a la fecha
            DateFrom = DateAdd("d", -1, DateFrom)
            
            ' Verificar si el día no es sábado (7) o domingo (1)
            If Weekday(DateFrom, vbMonday) <= 5 Then
                Days = Days - 1  ' Restar un día laborable
            End If
        Loop

        DateFrom = Format(DateFrom, "mm/dd/yy")
        DueDate = Format(CDate(MReport(fila, ColDueDate)), "mm/dd/yy")

        'Identificamos el delay
        If DueDate <= DateFrom Then hReport.Range("A" & fila + 1 & ":FC" & fila + 1).Font.Color = RGB(255, 0, 0)

      Else  
        hReport.Range("A" & fila + 1 & ":FC" & fila + 1).Font.Color = RGB(112, 48, 160)
      End If
    End If
  Next fila  

  SetOrder

End Sub


' ==============================================================
'                             EXTRAS
' ==============================================================

Sub SetDictionary(Matriz As Variant, Dict As Variant)
  For fila = 1 to UBound(Matriz)
    Key = Matriz(fila, 1)
    value = Matriz(fila, 2)
    If Not Dict.Exists(key) Then Dict.Add key, value
  Next fila
End Sub

Sub SetDictionary2(Matriz As Variant, Dict As Variant)
  For fila = 3 To UBound(Matriz)  
    Account =  Trim(Matriz(fila, 1))                                          
    Set DAccountsTimes = CreateObject("Scripting.Dictionary") 

    For column = 2 To UBound(Matriz, 2)                                            
      Header = Trim(Matriz(1, column))
      Type_ = Trim(Matriz(2, column))
      key = Type_ & "|" & Header

      If Matriz(fila, column) <> "" Then
        DAccountsTimes.Add key, Matriz(fila, column)
      Else
        DAccountsTimes.Add key, "Unknow"
      End If
    Next column

    If Not Dict.Exists(Account) Then Dict.Add Account, DAccountsTimes
  Next fila
End Sub

Sub SetDictionary3(Matriz As Variant, Dict As Variant)
  For column = 2 To UBound(Matriz, 2)                                            
    Header = Matriz(1, column)
    Type_ = Matriz(2, column)
    If Not Dict.Exists(Type_) Then 
      Dict.Add Type_, Header
    Else  
      Dict.Add Type_ & "_", Header
    End if
  Next column    
End Sub

Sub ApplyFormat(hSave As Variant)
  Dim rango As Range

  LastRow = hSave.Cells(hSave.Rows.Count, 1).End(xlUp).Row
  LCN = hSave.Cells(1, hSave.Columns.Count).End(xlToLeft).Column
  LCL = Split(Cells(1, LCN).Address, "$")(1)

  Set rango = hSave.Range("A2:" & LCL & LastRow)

  ' Aplicar bordes delgados en todas las celdas del rango
  With rango.Borders
      .LineStyle = xlContinuous          ' Estilo de l?nea continua
      .Color = RGB(0, 0, 0)             ' Color negro
      .TintAndShade = 0                 ' Sin sombra
      .Weight = xlThin                  ' Bordes delgados
  End With

  ' Configurar la fuente en todo el rango
  With rango.Font
      .Name = "Calibri"                  ' Estilo de letra Calibri
      .Size = 9                          ' Tama?o de letra
      .Color = RGB(0, 0, 0)              ' Color de letra negro
  End With

  'Ajustar el ancho de las columnas
  hSave.Columns("A:" & LCL).AutoFit

  ' Centrar el texto horizontal y verticalmente
  With rango
      .HorizontalAlignment = xlCenter     ' Alineaci?n horizontal (centrado)
      .VerticalAlignment = xlCenter       ' Alineaci?n vertical (centrado)
  End With
End Sub

Sub SetOrder()
    Dim LastRow As Long
    Dim LCN As Long
    Dim LCL As String
    Dim auxCol As Long
    Dim i As Long
    Dim rango As Range

    ' Obtener la última fila y la última columna con datos
    LastRow = hReport.Cells(hReport.Rows.Count, 1).End(xlUp).Row
    LCN = hReport.Cells(1, hReport.Columns.Count).End(xlToLeft).Column
    LCL = Split(Cells(1, LCN).Address, "$")(1)

    ' Crear columna auxiliar en la columna siguiente al último dato
    auxCol = LCN + 1
    LCLaux = Split(Cells(1, auxCol).Address, "$")(1)

    ' Agregar los valores numéricos de color en la columna auxiliar
    For i = 2 To LastRow ' Comienza desde la fila 2 para evitar los encabezados
        If hReport.Cells(i, 1).Font.Color = RGB(255, 0, 0) Then
            hReport.Cells(i, auxCol).Value = 1  ' Rojo
        ElseIf hReport.Cells(i, 1).Font.Color = RGB(112, 48, 160) Then
            hReport.Cells(i, auxCol).Value = 2  ' Morado
        Else
            hReport.Cells(i, auxCol).Value = 3  ' Otros colores
        End If
    Next i

  Set rango = hReport.Range("A1:" & LCLaux & LastRow)

  With hReport.Sort
    .SortFields.Clear
    .SortFields.Add2 key:=Range(LCLaux & "2:" & LCLaux & totalFilas + 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange rango
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With

    ' Eliminar la columna auxiliar después de ordenar
    hReport.Columns(auxCol).Delete
End Sub

Function GetDicValue(key As Variant, Dict As Variant) As Variant
  If Not Dict.Exists(key) Then 
    GetDicValue = "Not Found"
  Else 
    GetDicValue = Dict.Item(key)
  End If
End Function 



