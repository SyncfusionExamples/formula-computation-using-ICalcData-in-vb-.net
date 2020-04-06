Imports Syncfusion.Calculate
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace CalculateDemo
    Class Program
        Private engine As CalcEngine = Nothing

        Public Shared Sub Main()
            Dim calcData As CalcData = New CalcData()
            calcData.SetValueRowCol1("10", 1, 1)
            calcData.SetValueRowCol1("20", 1, 2)
            Dim engine As CalcEngine = New CalcEngine(calcData)
            Dim formula As String = "SUM (A1,B1)"
            Dim result As String = engine.ParseAndComputeFormula(formula)
            Console.WriteLine("A1 = {0}", "10")
            Console.WriteLine("B1 = {0}", "20")
            Console.WriteLine("SUM (A1,B1) is {0}", result)
            Console.ReadLine()

        End Sub
    End Class

    Public Class CalcData
        Implements ICalcData

        Private values As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()
        Public Event ValueChanged As ValueChangedEventHandler

        Private Sub OnValueChanged(ByVal row As Integer, ByVal col As Integer, ByVal value As String)
            RaiseEvent ValueChanged(Me, New ValueChangedEventArgs(row, col, value))
        End Sub

        Public Function GetValueRowCol1(row As Integer, col As Integer) As Object Implements ICalcData.GetValueRowCol
            Dim value As Object = Nothing
            Dim key As String = RangeInfo.GetAlphaLabel(col) + row.ToString()
            Me.values.TryGetValue(key, value)
            Return value
        End Function

        Public Sub SetValueRowCol1(value As Object, row As Integer, col As Integer) Implements ICalcData.SetValueRowCol
            Dim key = RangeInfo.GetAlphaLabel(col) + row.ToString()

            If Not values.ContainsKey(key) Then
                values.Add(key, value)
            ElseIf values.ContainsKey(key) AndAlso values(key) <> value Then
                values(key) = value
            End If
        End Sub

        Public Event ValueChanged1(sender As Object, e As ValueChangedEventArgs) Implements ICalcData.ValueChanged

        Public Sub WireParentObject1() Implements ICalcData.WireParentObject

        End Sub
    End Class
End Namespace


