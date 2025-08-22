# Formula computation using ICalcData in vb .net

This example demonstrates how to compute the formulas using `ICalcData` in Vb.Net.

The [ICalcData](https://help.syncfusion.com/cr/windowsforms/Syncfusion.Calculate.ICalcData.html) object can be integrated into [CalcEngine](https://help.syncfusion.com/cr/windowsforms/Syncfusion.Calculate.CalcEngine.html) by passing it through constructor. Now, you can compute the expressions or equations using CalcEngine.

The [ParseAndComputeFormula](https://help.syncfusion.com/cr/windowsforms/Syncfusion.Calculate.CalcEngine.html#Syncfusion_Calculate_CalcEngine_ParseAndComputeFormula_System_String_) method of [CalcEngine](https://help.syncfusion.com/cr/windowsforms/Syncfusion.Calculate.CalcEngine.html) is used to evaluate the formulas using the values from `ICalcData` object by cell references.

``` vb
Dim calcData As CalcData = New CalcData()

Dim engine As CalcEngine = New CalcEngine(calcData)
            
Dim formula As String = "SUM (A1,B1)"
        
Dim result As String = engine.ParseAndComputeFormula(formula)
```