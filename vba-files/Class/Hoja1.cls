VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub bnt_Facturar_Click()

    Dim register, inventary As Object
    Dim item, separateData, sale As Variant
    Dim data As String
    Dim CanCurrent, CanSale As Integer

    Application.ScreenUpdating = False
    Worksheets("Forms").Range("$B$6") = Empty
    If Range("tbl_venta[[#headers],[Cliente]]").Offset(1, 0).Value = Empty Then
        Worksheets("Forms").Range("$B$6") = "No hay productos registrados para la venta"
        Exit Sub
    End If
    Set inventary = ThisWorkbook.Worksheets("Inventario").Range("tbl_inventarios[Detalle de Productos]")
    Set register = Range("tbl_venta[Producto]")
    data = Range("tbl_venta[[#headers],[Cliente]]").Offset(1, 0).Value
    separateData = VBA.Split(data, "-")
    ThisWorkbook.Worksheets(Trim(separateData(0))).Select
    ActiveSheet.Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    For Each item In register
        For Each sale In inventary
            If Trim(item) = Trim(sale) Then
                CanCurrent = sale.Offset(, 2)
                CanSale = sale.Offset(, 6)
                sale.Offset(, 2) = CanCurrent - item.Offset(, 2)
                sale.Offset(, 6) = CanSale + item.Offset(, 2)
                Exit For
            End If
        Next sale
        ActiveCell.Offset(, -2) = Date
        ActiveCell.Offset(, -1) = item.Offset(, 2).Value
        ActiveCell = item.Value
        ActiveCell.Offset(, 1) = CInt(item.Offset(, 4).Value) / CInt(item.Offset(, 2).Value)
        If Trim(separateData(1)) = "Pendiente" Then: ActiveCell.Offset(, 2) = item.Offset(, 4).Value
            If Trim(separateData(1)) = "Pagado" Then: ActiveCell.Offset(, 3) = "OK"
                ActiveCell.Offset(1, 0).Select
            Next item
            ThisWorkbook.Worksheets("Forms").Select
            ActiveSheet.Range("tbl_venta[#Data]").Select
            For item = 1 To register.Count
                Selection.ListObject.ListRows(1).Delete
                Next
                ActiveSheet.Range("tbl_venta[[#Headers],[Producto]]").Select
                ListUsuario.Value = Empty
                Application.ScreenUpdating = True
End Sub

Private Sub btn_Carrito_Click()

    Dim status As String

    Application.ScreenUpdating = False
    Worksheets("Forms").Range("$B$6") = Empty
    If ListaProductos.Value = Empty Or txt_Price.Value = Empty Or txt_Cantidad.Value = Empty Or ListUsuario.Value = Empty Then
        Worksheets("Forms").Range("$B$6") = "Hay Campos de venta vacios"
        Exit Sub
    End If

    Range("tbl_venta[[#Headers],[Producto]]").Select
    ActiveCell.Offset(1, 0).Select
    If ActiveCell <> Empty And ActiveCell.Offset(1, 0) <> "Total" Then: ActiveCell.End(xlDown).Offset(-1, 0).Select
        If ActiveCell <> Empty Then
            Selection.ListObject.ListRows.Add AlwaysInsert:=True
            ActiveCell.Offset(1, 0).Select
        End If

        If OptPagado.Value = True Then
            status = "Pagado"
        ElseIf OptPendiente.Value = True Then
            status = "Pendiente"
        End If

        ActiveCell = ListaProductos.Value
        ActiveCell.Offset(, 1) = txt_Price.Value
        ActiveCell.Offset(, 2) = txt_Cantidad.Value
        ActiveCell.Offset(, 3) = ListUsuario.Value & " - " & status
        ActiveCell.Offset(, 4) = CInt(txt_Price.Value) * CInt(txt_Cantidad.Value)

        ListaProductos.Value = Empty
        txt_Price.Value = Empty
        txt_Cantidad.Value = Empty
        OptPendiente.Value = True

        Application.ScreenUpdating = True

End Sub

Private Sub ListaProductos_Change()

    Dim sh As Worksheet
    Dim price As Range
    Dim item As Variant

    Set sh = ThisWorkbook.Worksheets("Inventario")
    Set price = sh.Range("tbl_inventarios[[Detalle de Productos]]")

    txt_Price.Value = Empty
    For Each item In price
        If Trim(item.Value) = Trim(ListaProductos.Value) Then
            txt_Price.Value = item.Offset(, 5).Value
            Exit For
        End If
    Next item

End Sub

Private Sub ListaProductos_GotFocus()

    Dim sh As Worksheet
    Dim product As Range
    Dim item As Variant

    Set sh = ThisWorkbook.Worksheets("Inventario")
    Set product = sh.Range("tbl_inventarios[[Detalle de Productos]]")

    Worksheets("Forms").Range("$B$6") = Empty
    ListaProductos.Clear
    For Each item In product
        ListaProductos.AddItem Trim(item)
    Next item

End Sub

Private Sub ListUsuario_GotFocus()

    Dim sh As Workbook
    Dim item As Variant

    Set sh = ThisWorkbook
    Worksheets("Forms").Range("$B$6") = Empty
    ListUsuario.Clear
    For Each item In sh.Worksheets
        If item.CodeName <> "Hoja1" And item.CodeName <> "Hoja11" And item.CodeName <> "Hoja2" Then: ListUsuario.AddItem Trim(item.Name)
        Next item

End Sub

Private Sub txt_Cantidad_Change()

    Dim sh As Worksheet
    Dim product As Range
    Dim item As Variant

    Set sh = ThisWorkbook.Worksheets("Inventario")
    Set product = sh.Range("tbl_inventarios[[Detalle de Productos]]")

    Worksheets("Forms").Range("$B$6") = Empty
    If txt_Cantidad.Value <> Empty Then
        For Each item In product
            If Trim(item.Value) = ListaProductos.Value And CInt(Trim(item.Offset(, 2).Value)) < CInt(Trim(txt_Cantidad.Value)) Then
                Worksheets("Forms").Range("$B$6") = "Disponibilidad del producto " & CStr(ListaProductos.Value) & " : " & CStr(item.Offset(, 2).Value)
                Exit For
            End If
        Next item
    End If

End Sub

Private Sub txt_Price_Change()
    Worksheets("Forms").Range("$B$6") = Empty
End Sub
