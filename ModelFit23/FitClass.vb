Imports System.Diagnostics
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports System.Xml

Namespace Gaskination
    Public Class FitClass
        Implements AmosGraphics.IPlugin
        'This plugin was written by John Lim July 2016 for James Gaskin

        Public Function Name() As String Implements AmosGraphics.IPlugin.Name
            Return "Model Fit Measures"
        End Function

        Public Function Description() As String Implements AmosGraphics.IPlugin.Description
            Return "Puts important measures of model fit into a table on an html document. See statwiki.kolobkreations.com for more information."
        End Function

        Public Function Mainsub() As Integer Implements AmosGraphics.IPlugin.MainSub
            'Fits the specified model.
            AmosGraphics.pd.AnalyzeCalculateEstimates()

            'Remove the old table files
            If (System.IO.File.Exists("ModelFit.html")) Then
                System.IO.File.Delete("ModelFit.html")
            End If

            'Start the amos debugger and create an object from the AmosEngine
            Dim debug As New AmosDebug.AmosDebug
            Dim Sem As New AmosEngineLib.AmosEngine
            Sem.NeedEstimates(SampleCorrelations)
            Sem.NeedEstimates(ImpliedCorrelations)

            'Get CFI estimate from xpath expression
            Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
            doc.Load(AmosGraphics.pd.ProjectName & ".AmosOutput")
            Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
            Dim e As Xml.XmlNode
            Dim eRoot As XmlElement = doc.DocumentElement
            e = eRoot.SelectSingleNode("body/div/div[@ntype='modelfit']/div[@nodecaption='Baseline Comparisons']/table/tbody/tr[position() = 1]/td[position() = 6]", nsmgr)
            Dim CFI As Double
            CFI = e.InnerText

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
            If CD > 5 Then
                sCD = "Terrible"
            ElseIf CD > 3 Then
                sCD = "Acceptable"
            Else
                sCD = "Excellent"
            End If
            If CFI > 0.95 Then
                sCFI = "Excellent"
            ElseIf CFI >= 0.9 Then
                sCFI = "Acceptable"
            Else
                sCFI = "Need More DF"
            End If
            If DTemp < 0.08 Then
                sSRMR = "Excellent"
            ElseIf DTemp <= 0.1 Then
                sSRMR = "Acceptable"
            Else
                sSRMR = "Terrible"
            End If
            If Sem.Rmsea < 0.06 Then
                sRMSEA = "Excellent"
            ElseIf Sem.Rmsea <= 0.08 Then
                sRMSEA = "Acceptable"
            Else
                sRMSEA = "Terrible"
            End If
            If Sem.Pclose > 0.05 Then
                sPclose = "Excellent"
            ElseIf Sem.Pclose > 0.01 Then
                sPclose = "Acceptable"
            Else
                sPclose = "Terrible"
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
            debug.PrintX("</td></tr></table><hr/>")

            'Write reference table and credits
            debug.PrintX("<h3> Cutoff Criteria*</h3><table><tr><th>Measure</th><th>Terrible</th><th>Acceptable</th><th>Excellent</th></tr>")
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
    End Class
End Namespace