Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports pdftron
Imports pdftron.PDF
Imports pdftron.Common
Imports System.Collections.Specialized

Public Class Form1
    Public FRMW As FRMW.FRMW
    Dim DOCU As DOCU.DOCU
    Dim convLog As ConversionLog.ConversionLog

    Dim swDOCU As StreamWriter
    Dim swSLCT As StreamWriter

    Dim overlayLogoDoc As PDFDoc
    Dim CurrentPage, overlayLogoPage As Page
    Dim clipAccountNumber, clipDocDate, clipNA, clipRA, clipQESP, clipDocType, clipInvoiceNumber, clipDueDate, clipTotalAmtDue, clipReturnAddress As New Rect
    Dim HelveticaRegularFont As PDF.Font

    Dim nameAddressList, remitAddressList, returnAddressList As New StringCollection

    Dim clientCode, CurrentPDFFileName, documentID, accountNumber, docDate, workDir, QESP, pieceID, prevPieceID, docType, invoiceNumber, dueDate, totalAmtDue As String
    Dim docNumber, currentPageNumber, origPageNumber, docPageCount, StartingPage, totalPages, pageTotal As Integer
    Dim selectBRE As Boolean
    Dim cancelledFlag As Boolean = False

    Structure TextAndStyle
        Public text As String
        Public fontName As String
        Public fontSIze As Double
    End Structure

    'BelWo Variables - to be removed prior to production
    Public OSG As Boolean = True
    Dim enc As Encoding = Encoding.Default
    Dim xmlOut As XmlTextWriter

    'Insertion of PDF images/objects (Remove if no images/messages are being added to the PDF)
    Dim overlayFrontDoc, overlayBackDoc As PDFDoc
    Dim overlayFrontPage, overlayBackPage As Page
    Dim XObjects As Dictionary(Of String, Element)

#Region "Form Events"

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim oRAngle As System.Drawing.Rectangle = New System.Drawing.Rectangle(0, 0, Me.Width, Me.Height)
        Dim oGradientBrush As Brush = New Drawing.Drawing2D.LinearGradientBrush(oRAngle, Color.WhiteSmoke, Color.Crimson, Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal)
        e.Graphics.FillRectangle(oGradientBrush, oRAngle)
    End Sub

    Private Sub Form1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        If (OSG) Then
            If (Environment.ExitCode = 0) And FRMW.parse("{NormalTermination}") <> "YES" Then
                cancelledFlag = True
                Throw New Exception("Program was cancelled while executing")
            End If
        Else
            cancelledFlag = True
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Interval = 1000
        Timer1.Enabled = True
        status("Starting")
    End Sub

    Private Sub status(ByVal txt As String)
        lblStatus.Text = txt
        Me.Refresh()
        Application.DoEvents()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        standardProcessing()
    End Sub

#End Region

#Region "OSG Process"

    Private Sub standardProcessing()
        Dim licenseKey As String = ""
        Dim docuFileName As String = ""
        Dim swSLCTFileName As String = ""

        FRMW = New FRMW.FRMW
        lblEXE.Text = Application.ExecutablePath
        FRMW.StandardInitialization(Application.ExecutablePath)
        convLog = New ConversionLog.ConversionLog("PDFextractDIVCTS")
        DOCU = New DOCU.DOCU
        If (OSG) Then
            FRMW.loadFrameworkApplicationConfigFile("PDFEXTRACT")
            licenseKey = FRMW.getJOBparameter("PDFTRONLICENSELEY")
            docuFileName = FRMW.getParameter("PDFEXTRACT.outputDOCUfile")
            CurrentPDFFileName = FRMW.getParameter("PDFEXTRACT.inputPDFfile")
            swSLCTFileName = FRMW.getParameter("PDFEXTRACT.OUTPUTSLCTFILE")
            clientCode = FRMW.getParameter("CLIENTCODE")
            workDir = FRMW.getParameter("WORKDIR")
        Else
            CurrentPDFFileName = "D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\DIVHRG-Commercial.pdf"
            clientCode = "DIVHRGComm"
            licenseKey = "OSG Billing Services(osgbilling.com):ENTCPU:1::W:AMS(20140622):8E4F78C23CAFD6B962824007400DD29C15AD420D33446C5017F1E6BEF5C7"
            docuFileName = "D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output\swdocu.txt"
            swSLCTFileName = "D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output\swSLCT.txt"
            workDir = "D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output"
            xmlOut = New XmlTextWriter("D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output\result.xml", enc)
            xmlOut.Formatting = Formatting.Indented
            xmlOut.WriteStartDocument()
            xmlOut.WriteStartElement("DOCS")
        End If

        swSLCT = New StreamWriter(swSLCTFileName, False, Encoding.Default)
        swDOCU = New StreamWriter(docuFileName, False, Encoding.Default)
        PDFNet.Initialize(licenseKey)

        SetParsingCoordinates()

        ProcessPDF()

        swDOCU.Flush() : swDOCU.Close()
        If (OSG) Then
            swSLCT.Flush() : swSLCT.Close()
        End If
        PDFNet.Terminate()
        If (OSG) Then
            convLog.ZIPandCopy()
            FRMW.StandardTermination()
        Else
            xmlOut.WriteEndDocument()
            xmlOut.Flush() : xmlOut.Close()
        End If
        Application.Exit()

    End Sub

    Private Sub ProcessPDF()
        ClearValues()

        'Open PDF file with PDFTron
        Using inDoc As New PDFDoc(CurrentPDFFileName)
            pageTotal = inDoc.GetPageCount

            'Load Fonts
            LoadFonts(inDoc)
            LoadOverlays()
            XObjects = New Dictionary(Of String, Element)
            AddImageXobjects(inDoc)

            While currentPageNumber < pageTotal
                currentPageNumber += 1 'Current page number will increment as blank pages and backer are added
                origPageNumber += 1

                CurrentPage = inDoc.GetPage(currentPageNumber)
                ProcessPage(inDoc)

            End While

            'Write DOCU record for last account
            documentID = Guid.NewGuid.ToString
            writeDOCUrecord(totalPages)
            status("Processing PDF page (" & origPageNumber.ToString & "); Saving Output PDF...")
            If (OSG) Then
                inDoc.Save(FRMW.getParameter("PDFExtract.OutputPDFFile"), SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)
            Else
                inDoc.Save("D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output\updated.pdf", SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)
            End If

        End Using

    End Sub

    Private Sub ClearValues()
        accountNumber = "" : docDate = "" : dueDate = ""
        nameAddressList = New StringCollection
        remitAddressList = New StringCollection
        returnAddressList = New StringCollection
        totalPages = 0 : docPageCount = 1
    End Sub

    Private Sub ProcessPage(inDoc As PDFDoc)
        Dim seq As String = ""

        QESP = GetPDFpageValue(clipQESP)
        If QESP.Contains(":") Then
            pieceID = QESP.Split(":"c)(2)
            seq = QESP.Split(":"c)(3)
        End If

        'Remove 2-D bar code
        WhiteOutContentBox(0.1, 8.0, 0.45, 0.8, , , , 1)

        If origPageNumber Mod 100 = 0 Then
            status("Processing PDF page (" & origPageNumber.ToString & ")")
        End If

        If pieceID <> prevPieceID Then
            If Integer.Parse(seq) = 1 Then
                'Start of document will have sequence number = 1
                ProcessPage1()
                prevPieceID = pieceID
            End If
        End If

        '  AdjustPagePosition(CurrentPage, -0.25, 0)

        prevPieceID = pieceID
        totalPages += 1
        docPageCount += 1

    End Sub

    Private Sub ProcessPage1()

        If docNumber > 0 Then
            'Write DOCU record
            documentID = Guid.NewGuid.ToString
            writeDOCUrecord(totalPages)
            ClearValues()
        End If

        'Get important values
        Dim tempColl As StringCollection = GetPDFpageValues(clipAccountNumber)
        accountNumber = tempColl(tempColl.Count - 1).ToString

        tempColl = GetPDFpageValues(clipDocDate)
        docDate = tempColl(tempColl.Count - 1).ToString

        tempColl = GetPDFpageValues(clipDueDate)
        dueDate = tempColl(tempColl.Count - 1).ToString

        tempColl = GetPDFpageValues(clipTotalAmtDue)
        totalAmtDue = tempColl(tempColl.Count - 1).ToString

        'docType = GetPDFpageValue(clipDocType)

        returnAddressList = GetPDFpageValues(clipReturnAddress)
        nameAddressList = GetPDFpageValues(clipNA)
        StartingPage = currentPageNumber
        documentID = Guid.NewGuid.ToString
        CreateSLCTentry()

        'Check values
        Dim tempDate As Date, tempDec As Decimal
        If accountNumber = "" Then Throw New Exception(convLog.addError("Account number not found", accountNumber, "123456789", "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        'If docType = "" Then Throw New Exception(convLog.addError("Document type not found", docType, , "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        'If invoiceNumber = "" Then Throw New Exception(convLog.addError("Invoice number not found", invoiceNumber, , "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        If Not Date.TryParse(docDate, tempDate) Then Throw New Exception(convLog.addError("Could not parse document date", docDate, "01/01/2016", "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        If Not Date.TryParse(dueDate, tempDate) Then Throw New Exception(convLog.addError("Could not parse due date", dueDate, "05/01/2016", "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        If Not Decimal.TryParse(totalAmtDue.Replace("$", ""), tempDec) Then Throw New Exception(convLog.addError("Could not parse total due amount", totalAmtDue, "$123.45", "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        If nameAddressList.Count = 0 Then Throw New Exception(convLog.addError("No name and address found", , , "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))
        If returnAddressList.Count = 0 Then Throw New Exception(convLog.addError("No return address found", , , "File: " & Path.GetFileName(CurrentPDFFileName) & " " & "Page " & origPageNumber))

        'White Out Address Box
        WhiteOutContentBox(0.5, 8, 4, 1.25, , , , 1)
        AddImage(CurrentPage, "FRONT", 0, 0)

        'Add logo and return address
        'AddImage(CurrentPage, "LOGO", 0.5, 9.45)
        'WriteOutText("Cintas", 0.525, 10.25,, 8, True)
        'WriteOutText("5600 West 73rd Street", 0.525, 10.14,, 8, True)
        'WriteOutText("Chicago, IL 60638", 0.525, 10.03,, 8, True)

        'WriteOutText("FOWARDING SERVICE REQUESTED", 1, 9.85, , 8)

        docNumber += 1
    End Sub

    Private Sub CreateSLCTentry()
        Dim SLCT As New SLCT.SLCT
        SLCT.documentId = documentID
        SLCT.applicationCode = FRMW.getParameter("SCSapplicationCode")
        SLCT.accountNumber = accountNumber
        SLCT.target = ""
        SLCT.addValue("insert1", QESP.Split(":"c)(4).Substring(1, 1))
        SLCT.addValue("insert2", QESP.Split(":"c)(4).Substring(2, 1))
        SLCT.addValue("insert3", QESP.Split(":"c)(4).Substring(3, 1))
        swSLCT.WriteLine(SLCT.SLCTrecord())
        SLCT = Nothing

    End Sub

#End Region

#Region "Standard PDF Procedures"

    Private Sub SetParsingCoordinates()

        clipAccountNumber.x1 = I2P(5.63)
        clipAccountNumber.y1 = I2P(9.08)
        clipAccountNumber.x2 = (clipAccountNumber.x1 + I2P(1.09))
        clipAccountNumber.y2 = (clipAccountNumber.y1 + I2P(0.15))

        'clipInvoiceNumber.x1 = I2P(6.3)
        'clipInvoiceNumber.y1 = I2P(10.15)
        'clipInvoiceNumber.x2 = (clipInvoiceNumber.x1 + I2P(0.75))
        'clipInvoiceNumber.y2 = (clipInvoiceNumber.y1 + I2P(0.3))

        clipDocDate.x1 = I2P(5.54)
        clipDocDate.y1 = I2P(8.94)
        clipDocDate.x2 = (clipDocDate.x1 + I2P(0.74))
        clipDocDate.y2 = (clipDocDate.y1 + I2P(0.15))

        clipDueDate.x1 = I2P(6.73)
        clipDueDate.y1 = I2P(2.69)
        clipDueDate.x2 = (clipDueDate.x1 + I2P(1.07))
        clipDueDate.y2 = (clipDueDate.y1 + I2P(0.19))

        clipTotalAmtDue.x1 = I2P(6.73)
        clipTotalAmtDue.y1 = I2P(2.46)
        clipTotalAmtDue.x2 = (clipTotalAmtDue.x1 + I2P(1.17))
        clipTotalAmtDue.y2 = (clipTotalAmtDue.y1 + I2P(0.2))


        clipReturnAddress.x1 = I2P(4.84)
        clipReturnAddress.y1 = I2P(1.25)
        clipReturnAddress.x2 = (clipReturnAddress.x1 + I2P(3.11))
        clipReturnAddress.y2 = (clipReturnAddress.y1 + I2P(0.48))

        'clipDocType.x1 = I2P(5.5)
        'clipDocType.y1 = I2P(10.5)
        'clipDocType.x2 = (clipDocType.x1 + I2P(3))
        'clipDocType.y2 = (clipDocType.y1 + I2P(0.5))

        clipNA.x1 = I2P(0.6)
        clipNA.y1 = I2P(8)
        clipNA.x2 = (clipNA.x1 + I2P(3.75))
        clipNA.y2 = (clipNA.y1 + I2P(1.25))

        'QESP Line
        clipQESP.x1 = I2P(0.65)
        clipQESP.y1 = I2P(0.13)
        clipQESP.x2 = (clipQESP.x1 + I2P(2.06))
        clipQESP.y2 = (clipQESP.y1 + I2P(0.2))

        CreateCropPage()

    End Sub

    Private Sub CreateCropPage()

        Dim cropDoc As New PDFDoc()
        Dim cropPDF As String = "" '
        If (OSG) Then
            cropPDF = FRMW.getParameter("WORKDIR") & "\crop.pdf"
        Else
            cropPDF = "D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output" & "\crop.pdf"
        End If

        If (OSG) Then
            cropDoc = New PDFDoc()
        Else
            cropDoc = New PDFDoc("D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\DIVHRG-Commercial.pdf")
        End If

        cropDoc.Save(cropPDF, SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)


        Dim page As Page = cropDoc.PageCreate(New Rect(0, 0, 612, 792))
        cropDoc.PageInsert(cropDoc.GetPageIterator(0), page)
        page = cropDoc.GetPage(1)

        'Remove x1 value from x2 for crop box creation
        CreateCropBox("ACCOUNT NUMBER", clipAccountNumber.x1, clipAccountNumber.y1, (clipAccountNumber.x2 - clipAccountNumber.x1), (clipAccountNumber.y2 - clipAccountNumber.y1), page, cropDoc)
        CreateCropBox("DOCUMENT DATE", clipDocDate.x1, clipDocDate.y1, (clipDocDate.x2 - clipDocDate.x1), (clipDocDate.y2 - clipDocDate.y1), page, cropDoc)
        CreateCropBox("DUE DATE", clipDueDate.x1, clipDueDate.y1, (clipDueDate.x2 - clipDueDate.x1), (clipDueDate.y2 - clipDueDate.y1), page, cropDoc)
        CreateCropBox("AMOUNT DUE", clipTotalAmtDue.x1, clipTotalAmtDue.y1, (clipTotalAmtDue.x2 - clipTotalAmtDue.x1), (clipTotalAmtDue.y2 - clipTotalAmtDue.y1), page, cropDoc)
        'CreateCropBox("DOC TYPE", clipDocType.x1, clipDocType.y1, (clipDocType.x2 - clipDocType.x1), (clipDocType.y2 - clipDocType.y1), page, cropDoc)
        'CreateCropBox("INVOICE NUMBER", clipInvoiceNumber.x1, clipInvoiceNumber.y1, (clipInvoiceNumber.x2 - clipInvoiceNumber.x1), (clipInvoiceNumber.y2 - clipInvoiceNumber.y1), page, cropDoc)
        CreateCropBox("RETURN ADDRESS", clipReturnAddress.x1, clipReturnAddress.y1, (clipReturnAddress.x2 - clipReturnAddress.x1), (clipReturnAddress.y2 - clipReturnAddress.y1), page, cropDoc)
        CreateCropBox("NAME & ADDRESS", clipNA.x1, clipNA.y1, (clipNA.x2 - clipNA.x1), (clipNA.y2 - clipNA.y1), page, cropDoc)
        CreateCropBox("QESP STRING", clipQESP.x1, clipQESP.y1, (clipQESP.x2 - clipQESP.x1), (clipQESP.y2 - clipQESP.y1), page, cropDoc)

        cropDoc.Save(cropPDF, SDF.SDFDoc.SaveOptions.e_compatibility + SDF.SDFDoc.SaveOptions.e_remove_unused)
        cropDoc.Close()

    End Sub

    Private Sub CreateCropBox(ByVal labelValue As String, ByVal x1Val As Double, ByVal y1Val As Double, ByVal x2Val As Double, ByVal y2Val As Double, ByVal PDFpage As Page, cropDoc As PDFDoc, Optional color1 As Double = 0.75, Optional color2 As Double = 0.75, Optional color3 As Double = 0.75, Optional opac As Double = 0.5)

        Dim elmBuilder As New ElementBuilder
        Dim elmWriter As New ElementWriter
        Dim element As Element
        elmWriter.Begin(PDFpage)
        elmBuilder.Reset() : elmBuilder.PathBegin()

        'Set crop box
        elmBuilder.CreateRect(x1Val, y1Val, x2Val, y2Val)
        elmBuilder.ClosePath()

        element = elmBuilder.PathEnd()
        element.SetPathFill(True)

        Dim gState As GState = element.GetGState
        gState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        gState.SetFillColor(New ColorPt(color1, color2, color3)) 'Default is gray
        gState.SetFillOpacity(opac)
        elmWriter.WriteElement(element)

        'Set text
        element = elmBuilder.CreateTextBegin(PDF.Font.Create(cropDoc, PDF.Font.StandardType1Font.e_helvetica_oblique, True), 8)
        element.GetGState.SetTextRenderMode(GState.TextRenderingMode.e_fill_text)
        element.GetGState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        element.GetGState.SetFillColor(New ColorPt(0, 0, 0))
        elmWriter.WriteElement(element)
        element = elmBuilder.CreateTextRun(labelValue)
        element.SetTextMatrix(1, 0, 0, 1, x1Val, (y1Val - 8))
        elmWriter.WriteElement(element)
        elmWriter.WriteElement(elmBuilder.CreateTextEnd())

        elmWriter.End()

    End Sub

    Private Sub LoadFonts(doc As PDFDoc)
        HelveticaRegularFont = pdftron.PDF.Font.Create(doc, PDF.Font.StandardType1Font.e_helvetica, False)
    End Sub

    Private Sub LoadOverlays()

        'Copy files over to use from local path rather than network

        'PDF Logo
        If (OSG) Then
            File.Copy(FRMW.getParameter("FRAMEWORKCLIENTPDFRESOURCESDIR") & "\HAR-012 j68253 Hargray LH.pdf", workDir & "\DIVHRG-Commerical.PDF", True)
            'File.Copy(FRMW.getParameter("FRAMEWORKCLIENTPDFRESOURCESDIR") & "\DIVCOH-BACK.PDF", workDir & "\DIVCOH-BACK.PDF", True)
        Else
            File.Copy("D:\OSGMigration\Git\PDFextractDIVHRG-Commerical" & "\HAR-012 j68253 Hargray LH.pdf", workDir & "\DIVHRG-Commerical.PDF", True)
            ' File.Copy("D:\OSGMigration\Git\PDFextractDIVHRG-Commerical\Output" & "\DIVCOH-BACK.PDF", workDir & "\DIVCOH-BACK.PDF", True)
        End If

        overlayFrontDoc = New PDFDoc(workDir & "\DIVHRG-Commerical.PDF")
        overlayFrontPage = overlayFrontDoc.GetPage(1)

        'overlayBackDoc = New PDFDoc(workDir & "\DIVCOH-BACK.PDF")
        'overlayBackPage = overlayBackDoc.GetPage(1)

    End Sub

    Private Sub AddImageXobjects(doc As PDFDoc)

        'Add xObjects of images to be reused throughout the output doc
        Dim EB As ElementBuilder

        EB = New ElementBuilder
        XObjects.Add("FRONT", EB.CreateForm(overlayFrontPage, doc))

        'EB = New ElementBuilder
        'XObjects.Add("BACK", EB.CreateForm(overlayBackPage, doc))

    End Sub

    Private Function GetPDFpageValue(clipRect As Rect) As String

        Dim docXML As New XmlDocument
        Dim X, Y, prevY As Double
        Dim x1Content As Double = clipRect.x1
        Dim y1Content As Double = clipRect.y1
        Dim x2Content As Double = clipRect.x2
        Dim y2Content As Double = clipRect.y2
        Dim contentValue As String = ""

        Using txt As TextExtractor = New TextExtractor
            Dim txtXML As String
            txt.Begin(CurrentPage, clipRect)
            txtXML = txt.GetAsXML(TextExtractor.XMLOutputFlags.e_output_bbox)
            docXML.LoadXml(txtXML)

            Dim tempRoot As XmlElement = docXML.DocumentElement
            Dim tempxnl1 As XmlNodeList
            tempxnl1 = Nothing
            tempxnl1 = tempRoot.SelectNodes("Flow/Para/Line")
            prevY = 0
            For Each elmC As XmlElement In tempxnl1
                Dim pos() As String = elmC.GetAttribute("box").Split(","c)
                X = pos(0) : Y = pos(1)

                'Page(Content)
                If (X >= x1Content) And (Y >= y1Content) And (X <= x2Content) And (Y <= y2Content) Then
                    If contentValue = "" Then
                        If prevY <> Math.Round(Y, 3) Then
                            contentValue = elmC.InnerText.Replace(vbLf, "")
                        End If
                    Else
                        contentValue = contentValue & elmC.InnerText.Replace(vbLf, "")
                    End If
                End If

                prevY = Math.Round(Y, 3)
                elmC = Nothing
            Next
        End Using

        Return contentValue

    End Function

    Private Function GetPDFpageValues(clipRect As Rect) As StringCollection
        Dim docXML As New XmlDocument
        Dim X, Y, prevY As Double
        Dim x1Content As Double = clipRect.x1
        Dim y1Content As Double = clipRect.y1
        Dim x2Content As Double = clipRect.x2
        Dim y2Content As Double = clipRect.y2
        Dim Values As New StringCollection

        Using txt As TextExtractor = New TextExtractor
            Dim txtXML As String
            txt.Begin(CurrentPage, clipRect)
            txtXML = txt.GetAsXML(TextExtractor.XMLOutputFlags.e_output_bbox)
            docXML.LoadXml(txtXML)

            Dim tempRoot As XmlElement = docXML.DocumentElement
            Dim tempxnl1 As XmlNodeList
            tempxnl1 = Nothing
            tempxnl1 = tempRoot.SelectNodes("Flow/Para/Line")
            prevY = 0
            For Each elmC As XmlElement In tempxnl1
                Dim pos() As String = elmC.GetAttribute("box").Split(","c)
                X = pos(0) : Y = pos(1)

                If (X >= x1Content) And (Y >= y1Content) And (X <= x2Content) And (Y <= y2Content) Then
                    If prevY <> Math.Round(Y, 3) Then
                        Values.Add(elmC.InnerText.Replace(vbLf, ""))
                    Else
                        Values(Values.Count - 1) = Values(Values.Count - 1) & elmC.InnerText.Replace(vbLf, "")
                    End If
                End If

                prevY = Math.Round(Y, 3)
                elmC = Nothing
            Next
        End Using

        Return Values
    End Function

    Private Function I2P(i As Decimal) As Decimal
        Return (i * 72)
    End Function

    Private Function CollectionToArray(Collection As StringCollection, ArraySize As Integer) As String()
        Dim Values(ArraySize) As String
        Dim i As Integer = 0
        For i = 0 To ArraySize
            If i <= Collection.Count - 1 Then
                Values(i) = Collection(i)
            Else
                Values(i) = ""
            End If
        Next
        Return Values
    End Function

    Private Sub WriteOutText(textToWrite As String, xPosition As Double, yPosition As Double, Optional fontType As String = "REGULAR",
                             Optional fontSize As Double = 10, Optional blue As Boolean = False, Optional alignment As String = "L")
        Dim eb As New ElementBuilder
        Dim writer As New ElementWriter
        Dim element As Element
        writer.Begin(CurrentPage)
        eb.Reset() : eb.PathBegin()

        element = eb.CreateTextBegin()
        element.GetGState.SetTextRenderMode(GState.TextRenderingMode.e_fill_text)
        element.GetGState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        Dim colors As New ColorPt(0, 0, 0)
        If blue Then colors.Set(0.05, 0.26, 0.52)
        element.GetGState.SetFillColor(colors)
        writer.WriteElement(element)

        Select Case fontType.ToUpper
            Case "REGULAR"
                'Helvetica
                element = eb.CreateTextRun(textToWrite, HelveticaRegularFont, fontSize)
            Case Else
                Throw New Exception(convLog.addError("Incorrect font type used in code, have tech take a look.", fontType.ToUpper, "REGULAR", , , , , True))
        End Select

        Select Case alignment
            Case "C"
                element.SetTextMatrix(1, 0, 0, 1, I2P(xPosition) - (element.GetTextLength / 2), I2P(yPosition))
            Case "R"
                element.SetTextMatrix(1, 0, 0, 1, I2P(xPosition) - element.GetTextLength, I2P(yPosition))
            Case Else
                element.SetTextMatrix(1, 0, 0, 1, I2P(xPosition), I2P(yPosition))
        End Select
        writer.WriteElement(element)
        writer.WriteElement(eb.CreateTextEnd())
        writer.End()

    End Sub

    Private Sub WhiteOutContentBox(x1Val As Double, y1Val As Double, x2Val As Double, y2Val As Double, Optional color1 As Double = 255, Optional color2 As Double = 255, Optional color3 As Double = 255, Optional opac As Double = 0.5)

        Dim elmBuilder As New ElementBuilder
        Dim elmWriter As New ElementWriter
        Dim element As Element
        elmWriter.Begin(CurrentPage)
        elmBuilder.Reset() : elmBuilder.PathBegin()

        'Set crop box
        elmBuilder.CreateRect(I2P(x1Val), I2P(y1Val), I2P(x2Val), I2P(y2Val))
        elmBuilder.ClosePath()

        element = elmBuilder.PathEnd()
        element.SetPathFill(True)

        Dim gState As GState = element.GetGState
        gState.SetFillColorSpace(ColorSpace.CreateDeviceRGB())
        gState.SetFillColor(New ColorPt(color1, color2, color3)) 'default color is white
        gState.SetFillOpacity(opac)
        elmWriter.WriteElement(element)

        elmWriter.End()

    End Sub

    Private Sub AddImage(PDFpage As Page, imageKey As String, xPosition As Double, yPosition As Double)
        Dim element As Element
        Dim EW As ElementWriter

        If Not XObjects.ContainsKey(imageKey) Then Throw New Exception("Invalid image key passed in code, have tech take a look.")
        element = XObjects(imageKey)

        EW = New ElementWriter
        EW.Begin(PDFpage, ElementWriter.WriteMode.e_underlay)
        element.GetGState().SetTransform(1, 0, 0, 1, I2P(xPosition), I2P(yPosition))
        'element.GetGState().SetTransform(0, 0, 0, 0, I2P(xPosition), I2P(yPosition))
        EW.WritePlacedElement(element)
        EW.End()
    End Sub

    Private Sub AdjustPagePosition(PDFpage As Page, xPosition As Double, yPosition As Double)
        Dim element As Element
        Dim EW As ElementWriter
        Dim builder As ElementBuilder = New ElementBuilder
        element = builder.CreateForm(PDFpage)
        EW = New ElementWriter
        EW.Begin(PDFpage, ElementWriter.WriteMode.e_replacement)
        element.GetGState().SetTransform(1, 0, 0, 1, I2P(xPosition), I2P(yPosition))
        EW.WritePlacedElement(element)
        EW.End()
    End Sub

#End Region

#Region "Global Functions/Routines"

    Private Sub writeDOCUrecord(totalPages As Integer)
        DOCU.Clear()
        DOCU.AccountNumber = accountNumber
        DOCU.DocumentID = documentID
        DOCU.ClientCode = clientCode
        DOCU.DocumentDate = fmtDate(docDate, "yyyy/MM/dd")
        DOCU.DocumentDueDate = fmtDate(dueDate, "yyyy/MM/dd")
        DOCU.DocumentType = "" 'docType
        DOCU.DocumentKey = ""
        DOCU.Print_StartPage = StartingPage
        DOCU.Print_NumberOfPages = totalPages
        DOCU.DeliveryIMBserviceType = "082"
        DOCU.AmountDue = totalAmtDue
        DOCU.MailingID = Strings.Right(nameAddressList(0), 9)
        'DOCU.MailingID = Strings.Right(nameAddressList(0), 9)
        'Name/Address Info
        nameAddressList(0) = "" 'Set mailing ID to blank
        removeLastAddressLine(nameAddressList) 'Remove IMB data
        DOCU.setOriginalAddress(CollectionToArray(nameAddressList, 5), 1, False)


        If (OSG) Then
            swDOCU.WriteLine(DOCU.GetXML)
        Else
            CreateDocumentNode()
        End If

    End Sub

    Public Sub CreateDocumentNode()
        xmlOut.WriteStartElement("DOC") '<DOC>
        xmlOut.WriteAttributeString("documentId", DOCU.DocumentID)
        xmlOut.WriteStartElement("DOCUMENT")
        xmlOut.WriteAttributeString("documentId", DOCU.DocumentID)
        xmlOut.WriteAttributeString("clientCode", DOCU.ClientCode)
        xmlOut.WriteAttributeString("billerCode", "")
        xmlOut.WriteAttributeString("merchantCode", "")
        xmlOut.WriteAttributeString("accountNumber", DOCU.AccountNumber)
        xmlOut.WriteAttributeString("customerNumber", DOCU.CustomerNumber)
        xmlOut.WriteAttributeString("documentType", "Statement")
        xmlOut.WriteAttributeString("documentKey", "")
        xmlOut.WriteAttributeString("documentDate", DOCU.DocumentDate)
        xmlOut.WriteAttributeString("documentDueDate", DOCU.DocumentDueDate)
        xmlOut.WriteAttributeString("amountDue", DOCU.AmountDue)
        xmlOut.WriteAttributeString("enrollmentToken", DOCU.EnrollmentToken)
        xmlOut.WriteAttributeString("enrollmentToken2", DOCU.EnrollmentToken2)
        xmlOut.WriteAttributeString("sequenceNumber", "")
        xmlOut.WriteAttributeString("externalDocumentId", DOCU.ExternalDocumentID)
        xmlOut.WriteAttributeString("DPBC", "")
        xmlOut.WriteAttributeString("IMB", "")
        xmlOut.WriteAttributeString("address1", DOCU.OriginalAddress(0))
        xmlOut.WriteAttributeString("address2", DOCU.OriginalAddress(1))
        xmlOut.WriteAttributeString("address3", DOCU.OriginalAddress(2))
        xmlOut.WriteAttributeString("address4", DOCU.OriginalAddress(3))
        xmlOut.WriteAttributeString("address5", DOCU.OriginalAddress(4))
        xmlOut.WriteAttributeString("address6", DOCU.OriginalAddress(5))
        xmlOut.WriteAttributeString("remitDPBC", DOCU.RemitDPBC)
        xmlOut.WriteAttributeString("remitIMB", "")
        xmlOut.WriteAttributeString("remitAddress1", DOCU.RemitAddress(0))
        xmlOut.WriteAttributeString("remitAddress2", DOCU.RemitAddress(1))
        xmlOut.WriteAttributeString("remitAddress3", DOCU.RemitAddress(2))
        xmlOut.WriteAttributeString("remitAddress4", DOCU.RemitAddress(3))
        xmlOut.WriteAttributeString("mailClass", "")
        xmlOut.WriteAttributeString("handlingCode", "")
        xmlOut.WriteAttributeString("locationCode", "")
        xmlOut.WriteAttributeString("groupCode1", "")
        xmlOut.WriteAttributeString("groupCode2", "")
        xmlOut.WriteAttributeString("groupCode3", "")
        xmlOut.WriteAttributeString("UDF1", "")
        xmlOut.WriteAttributeString("UDF2", "")
        xmlOut.WriteAttributeString("UDF3", "")
        xmlOut.WriteAttributeString("UDF4", "")
        xmlOut.WriteAttributeString("UDF5", "")
        xmlOut.WriteAttributeString("PrintNumberOfPages", DOCU.Print_NumberOfPages)
        xmlOut.WriteEndElement()
        xmlOut.WriteEndElement() '</DOC> 
    End Sub

    Private Sub removeLastAddressLine(addressList As StringCollection)
        Dim addressLine As String = addressList(addressList.Count - 1)
        addressLine = addressLine.Replace("A", "").Replace("D", "").Replace("F", "").Replace("T", "")
        If addressLine.Trim = "" Then
            addressList(addressList.Count - 1) = ""
        End If
    End Sub

    Private Function fmtAmount(inputAmount As String, Optional displayCurrencySign As Boolean = True, Optional numberOfDecimalDigits As Integer = 2, Optional impliedDigit As Integer = 0, Optional validate As Boolean = False) As String

        Dim dDec As Decimal

        If Decimal.TryParse(inputAmount, dDec) Then
            dDec = inputAmount
            dDec /= (10 ^ impliedDigit)

            If displayCurrencySign Then
                inputAmount = FormatCurrency(dDec, numberOfDecimalDigits, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.False, Microsoft.VisualBasic.TriState.True)
            Else
                inputAmount = FormatNumber(dDec, numberOfDecimalDigits, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.False, Microsoft.VisualBasic.TriState.True)
            End If
        Else
            If validate Then
                Throw New Exception(convLog.addError("Not a valid amount value", inputAmount, , accountNumber, , docNumber))
            End If
        End If

        Return inputAmount

    End Function

    Private Function fmtDate(inputDate As String, Optional formatToUse As String = "MM/dd/yy", Optional validate As Boolean = False) As String

        Dim dateToParse As Date

        If Date.TryParse(inputDate, dateToParse) Then
            inputDate = Date.Parse(inputDate).ToString(formatToUse)
        Else
            If validate Then
                Throw New Exception(convLog.addError("Not a valid data value", inputDate, , accountNumber, , docNumber))
            End If
        End If

        Return inputDate

    End Function

#End Region

End Class

