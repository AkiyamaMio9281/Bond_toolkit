Attribute VB_Name = "RunScenariosModule"
Option Explicit

' Bond metrics returned as Variant array:
'   (1) Price
'   (2) Macaulay Duration (years)
'   (3) Modified Duration (years)  w.r.t nominal y = j^(m)
'   (4) Convexity (years^2)        w.r.t nominal y = j^(m)

Private Function BondCalc(ByVal FV As Double, ByVal c As Double, ByVal T As Double, _
                          ByVal p As Long, ByVal j As Double, ByVal m As Long) As Variant
    Dim N As Long
    N = CLng(T * p + 0.0000001)

    Dim coupon As Double
    coupon = FV * c / p

    Dim a As Double
    a = 1# + j / m

    Dim k As Long
    Dim t As Double, df As Double, cf As Double, pv As Double
    Dim price As Double, wsum As Double, convsum As Double

    For k = 1 To N
        t = k / p

        cf = coupon
        If k = N Then cf = coupon + FV

        df = a ^ (-m * t)
        pv = cf * df

        price = price + pv
        wsum = wsum + t * pv
        convsum = convsum + t * (t + 1# / m) * pv
    Next k

    Dim out(1 To 4) As Double
    out(1) = price

    If price <> 0# Then
        out(2) = wsum / price                 ' Macaulay
        out(3) = out(2) / a                   ' Modified = Mac / (1 + j/m)
        out(4) = convsum / (price * a ^ 2)    ' Convexity
    Else
        out(2) = 0#: out(3) = 0#: out(4) = 0#
    End If

    BondCalc = out
End Function

Public Sub RunScenarios()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Bond Toolkit")

    ' Read base inputs
    Dim FV As Double, c As Double, T As Double
    Dim p As Long, j As Double, m As Long

    FV = ws.Range("B5").Value
    c = ws.Range("B6").Value
    T = ws.Range("B7").Value
    p = CLng(ws.Range("B8").Value)
    j = ws.Range("B9").Value
    m = CLng(ws.Range("B10").Value)

    Dim base As Variant
    base = BondCalc(FV, c, T, p, j, m)

    ' Find scenario block
    Dim topCell As Range
    Set topCell = ws.Columns(1).Find(What:="SCENARIO TESTING", LookAt:=xlPart)
    If topCell Is Nothing Then
        MsgBox "Cannot find scenario block.", vbExclamation
        Exit Sub
    End If

    Dim scenHeaderRow As Long
    scenHeaderRow = topCell.Row + 1

    Dim r As Long
    r = scenHeaderRow + 1

    Application.ScreenUpdating = False

    Do While ws.Cells(r, 1).Value <> ""
        Dim dy_bp As Double
        Dim altmVal As Variant
        Dim scen_m As Long
        Dim dy As Double

        dy_bp = ws.Cells(r, 2).Value
        altmVal = ws.Cells(r, 3).Value

        If IsEmpty(altmVal) Or altmVal = "" Then
            scen_m = m
        Else
            scen_m = CLng(altmVal)
        End If

        dy = dy_bp / 10000#     ' bp -> decimal

        Dim scen As Variant
        scen = BondCalc(FV, c, T, p, j + dy, scen_m)

        ' Actual price & pct change vs base
        ws.Cells(r, 4).Value = scen(1)
        If base(1) <> 0# Then
            ws.Cells(r, 5).Value = scen(1) / base(1) - 1#
        Else
            ws.Cells(r, 5).Value = ""
        End If

        ' Approximations only when compounding m matches base
        Dim estDur As Variant, estDurConv As Variant
        If scen_m = m Then
            estDur = base(1) * (1# - base(3) * dy)
            estDurConv = base(1) * (1# - base(3) * dy + 0.5 * base(4) * dy * dy)
        Else
            estDur = ""
            estDurConv = ""
        End If

        ws.Cells(r, 6).Value = estDur
        ws.Cells(r, 7).Value = estDurConv

        If estDur = "" Or scen(1) = 0# Then
            ws.Cells(r, 8).Value = ""
        Else
            ws.Cells(r, 8).Value = estDur / scen(1) - 1#
        End If

        If estDurConv = "" Or scen(1) = 0# Then
            ws.Cells(r, 9).Value = ""
        Else
            ws.Cells(r, 9).Value = estDurConv / scen(1) - 1#
        End If

        r = r + 1
    Loop

    ' Apply number formats to scenario output area
    ws.Range(ws.Cells(scenHeaderRow + 1, 4), ws.Cells(r - 1, 4)).NumberFormat = "$#,##0.00"
    ws.Range(ws.Cells(scenHeaderRow + 1, 6), ws.Cells(r - 1, 7)).NumberFormat = "$#,##0.00"
    ws.Range(ws.Cells(scenHeaderRow + 1, 5), ws.Cells(r - 1, 5)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(scenHeaderRow + 1, 8), ws.Cells(r - 1, 9)).NumberFormat = "0.00%"

    Application.ScreenUpdating = True

    MsgBox "Scenarios updated.", vbInformation
End Sub
