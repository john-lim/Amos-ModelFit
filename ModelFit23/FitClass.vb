#Region "Imports"
Imports System.Diagnostics
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports System.Xml
Imports AmosGraphics
#End Region

Public Class FitClass
    Implements AmosGraphics.IPlugin
    'This plugin was written by John Lim July 2016 for James Gaskin

#Region "Name Plugin"
    Public Function Name() As String Implements AmosGraphics.IPlugin.Name
        Return "Model Fit Measures"
    End Function

    Public Function Description() As String Implements AmosGraphics.IPlugin.Description
        Return "Puts important measures of model fit into a table on an html document. See statwiki.kolobkreations.com for more information."
    End Function
#End Region

    Public Function Mainsub() As Integer Implements AmosGraphics.IPlugin.MainSub
#Region "Setup Model"
        'Fits the specified model.
        pd.GetCheckBox("AnalysisPropertiesForm", "ModsCheck").Checked = True
        pd.GetCheckBox("AnalysisPropertiesForm", "ResidualMomCheck").Checked = True
        pd.AnalyzeCalculateEstimates()

        'Remove the old table files
        If (System.IO.File.Exists("ModelFit.html")) Then
            System.IO.File.Delete("ModelFit.html")
        End If

        'Start the amos debugger and create an object from the AmosEngine
        Dim debug As New AmosDebug.AmosDebug
        Dim Sem As New AmosEngineLib.AmosEngine
        Sem.NeedEstimates(SampleCorrelations)
        Sem.NeedEstimates(ImpliedCorrelations)
#End Region

        'Get CFI estimate from xpath expression
        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(AmosGraphics.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim e As Xml.XmlNode
        Dim eRoot As XmlElement = doc.DocumentElement
        e = eRoot.SelectSingleNode("body/div/div[@ntype='modelfit']/div[@nodecaption='Baseline Comparisons']/table/tbody/tr[position() = 1]/td[position() = 6]", nsmgr)
        Dim CFI As Double = e.InnerText
        Dim tableSRC As XmlElement 'Table of body values of SRC
        Dim headSRC As XmlElement 'Table of variables
        tableSRC = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='matrices']/div[@ntype='ppml'][position() = 2]/table/tbody", nsmgr)
        headSRC = eRoot.SelectSingleNode("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='matrices']/div[@ntype='ppml'][position() = 2]/table/thead", nsmgr)

        Dim iObserved As Integer
        For Each a As PDElement In pd.PDElements
            If a.IsObservedVariable Then 'Checks if the variable is latent
                iObserved += 1 'Will return the number of variables in the model.
            End If
        Next

        Dim iCount As Integer = 0
        Dim iCount2 As Integer = iObserved
        Dim iCount3 As Integer = 0
        Dim iCount4 As Integer = 1
        Dim dictSRC As New Dictionary(Of Double, String)()
        Dim dSum As Double
        Dim listValues As New List(Of varSummed)

        For a = 0 To iObserved - 1
            For b = 1 To iCount2
                dSum = dSum + Math.Abs(MatrixElement(tableSRC, (b + iCount3), (a + 1)))
            Next
            For c = 1 To iCount4
                dSum = dSum + Math.Abs(MatrixElement(tableSRC, (a + 1), c))
            Next
            iCount2 -= 1
            iCount3 += 1
            iCount4 += 1
            Dim oValues As New varSummed(MatrixName(headSRC, 1, (a + 1)), dSum)
            listValues.Add(oValues)
            dSum = 0
        Next
        listValues.Sort(Function(x, y) y.Total.CompareTo(x.Total))
        listValues = listValues.OrderBy(Function(x) x.Total).ToList()

        'Specify and fit the object to the model
        AmosGraphics.pd.SpecifyModel(Sem)
        Sem.FitModel()

        'Calculate SRMR
        Dim N As Integer
        Dim i As Integer
        Dim j As Integer
        Dim DTemp As Double
        Dim Sample(,) As Double
        Dim Implied(,) As Double
        Sem.GetEstimates(SampleCorrelations, Sample)
        Sem.GetEstimates(ImpliedCorrelations, Implied)
        N = UBound(Sample, 1) + 1
        DTemp = 0
        For i = 1 To N - 1
            For j = 0 To i - 1
                DTemp = DTemp + (Sample(i, j) - Implied(i, j)) ^ 2
            Next
        Next
        DTemp = System.Math.Sqrt(DTemp / (N * (N - 1) / 2))

        Dim CD As Double = Sem.Cmin / Sem.Df

        'Conditionals for interpretation column
        Dim sCD As String = ""
        Dim sCFI As String = ""
        Dim sSRMR As String = ""
        Dim sRMSEA As String = ""
        Dim sPclose As String = ""
        Dim iBad As Integer = 0
        Dim iGood As Integer = 0
        If CD > 5 Then
            sCD = "Terrible"
            iBad += 1
        ElseIf CD > 3 Then
            sCD = "Acceptable"
            iGood += 1
        Else
            sCD = "Excellent"
        End If
        If CFI > 0.95 Then
            sCFI = "Excellent"
        ElseIf CFI >= 0.9 Then
            sCFI = "Acceptable"
            iGood += 1
        Else
            sCFI = "Need More DF"
            iBad += 1
        End If
        If DTemp < 0.08 Then
            sSRMR = "Excellent"
        ElseIf DTemp <= 0.1 Then
            sSRMR = "Acceptable"
            iGood += 1
        Else
            sSRMR = "Terrible"
            iBad += 1
        End If
        If Sem.Rmsea < 0.06 Then
            sRMSEA = "Excellent"
        ElseIf Sem.Rmsea <= 0.08 Then
            sRMSEA = "Acceptable"
            iGood += 1
        Else
            sRMSEA = "Terrible"
            iBad += 1
        End If
        If Sem.Pclose > 0.05 Then
            sPclose = "Excellent"
        ElseIf Sem.Pclose > 0.01 Then
            sPclose = "Acceptable"
            iGood += 1
        Else
            sPclose = "Terrible"
            iBad += 1
        End If

        'Set up the listener To output the debugs
        Dim resultWriter As New TextWriterTraceListener("ModelFit.html")
        Trace.Listeners.Add(resultWriter)

        'Write the beginning Of the document
        debug.PrintX("<html><body><h1>Model Fit Measures</h1><hr/>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Measure</th><th>Estimate</th><th>Threshold</th><th>Interpretation</th></tr>")
        debug.PrintX("<tr><td>CMIN</td><td>")
        debug.PrintX(Sem.Cmin.ToString("#0.000"))
        debug.PrintX("</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>DF</td><td>")
        debug.PrintX(Sem.Df)
        debug.PrintX("</td><td>--</td><td>--</td></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>")
        debug.PrintX(CD.ToString("#0.000"))
        debug.PrintX("</td><td>Between 1 and 3</td><td>")
        debug.PrintX(sCD)
        debug.PrintX("</td></tr><tr><td>CFI</td><td>")
        debug.PrintX(CFI.ToString("#0.000"))
        debug.PrintX("</td><td>>0.95</td><td>")
        debug.PrintX(sCFI)
        debug.PrintX("</td></tr><tr><td>SRMR</td><td>")
        debug.PrintX(DTemp.ToString("#0.000"))
        debug.PrintX("</td><td><0.08</td><td>")
        debug.PrintX(sSRMR)
        debug.PrintX("</td></tr><tr><td>RMSEA</td><td>")
        debug.PrintX(Sem.Rmsea.ToString("#0.000"))
        debug.PrintX("</td><td><0.06</td><td>")
        debug.PrintX(sRMSEA)
        debug.PrintX("</td></tr><tr><td>PClose</td><td>")
        debug.PrintX(Sem.Pclose.ToString("#0.000"))
        debug.PrintX("</td><td>>0.05</td><td>")
        debug.PrintX(sPclose)
        debug.PrintX("</td></tr></table><br>")
        If iGood = 0 And iBad = 0 Then
            debug.PrintX("Congratulations, your model fit is excellent!")
        ElseIf iGood > 0 And iBad = 0 Then
            debug.PrintX("Congratulations, your model fit is acceptable.")
        Else

            debug.PrintX("Unfortunately, your model fit could improve. Based on the standardized residual covariances, we recommend removing " + listValues.First.Name + ".")
        End If

        'Write reference table and credits
        debug.PrintX("<hr/><h3> Cutoff Criteria*</h3><table><tr><th>Measure</th><th>Terrible</th><th>Acceptable</th><th>Excellent</th></tr>")
        debug.PrintX("<tr><td>CMIN/DF</td><td>> 5</td><td>> 3</td><td>> 1</td></tr>")
        debug.PrintX("</td></tr><tr><td>CFI</td><td><0.90</td><td><0.95</td><td>>0.95</td></tr>")
        debug.PrintX("</td></tr><tr><td>SRMR</td><td>>0.10</td><td>>0.08</td><td><0.08</td></tr>")
        debug.PrintX("</td></tr><tr><td>RMSEA</td><td>>0.08</td><td>>0.06</td><td><0.06</td></tr>")
        debug.PrintX("</td></tr><tr><td>PClose</td><td><0.01</td><td><0.05</td><td>>0.05</td></tr></table>")
        debug.PrintX("<p>*Note: Hu and Bentler (1999, ""Cutoff Criteria for Fit Indexes in Covariance Structure Analysis: Conventional Criteria Versus New Alternatives"") recommend combinations of measures. Personally, I prefer a combination of CFI>0.95 and SRMR<0.08. To further solidify evidence, add the RMSEA<0.06.</p>")
        debug.PrintX("<p>**If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J. & Lim, J. (2016), ""Model Fit Measures"", AMOS Plugin. <a href=\""http://statwiki.kolobkreations.com"">Gaskination's StatWiki</a>.</p>")

        'Write Style And close
        debug.PrintX("<style>h1{margin-left:60px;}table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Sem.Dispose()
        Process.Start("ModelFit.html")
        Return 0
    End Function

#Region "Matrix functions"
    'This function gets a value of type double from the xml matrix
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double
        Dim e As XmlElement
        e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
        MatrixElement = CDbl(e.GetAttribute("x"))
    End Function
    'This function gets a value of type string from the matrix
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String
        Dim e As XmlElement
        'The row is offset one.
        e = eTableBody.ChildNodes(row - 1).ChildNodes(column)
        MatrixName = e.InnerText
    End Function
#End Region
End Class

Public Class varSummed
    Public Name As String
    Public Total As Double

    Public Sub New(ByVal sName As String, ByVal dTotal As Double)
        'constructor
        Name = sName
        Total = dTotal
        'storing the values in constructor
    End Sub
End Class
