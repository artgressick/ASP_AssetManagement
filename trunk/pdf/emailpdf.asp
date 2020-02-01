<!-- #include file="../includes/openconn.asp" -->
<%
	'RRF - Set Variables for Active Tool Kit
	'RRF - MaxRowsPerPage is the number of Fields on the 4th Page the inventory page of the PDF
	MaxRowsPerPage = 70
	'RRF - iFieldFlag is a attribute code that make the field flatten and keep the imported data.
	iFieldFlag = -998
	'RRF - bDoFormFormatting will keep the text formats that is designated on the PDF fields
	bDoFormFormatting = True
	'RRF - sUFileName is the Name of the Final PDF file to send to the customer.  
	sUFileName = "test.pdf"
    
	'find the cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute ViewInvoicebyCart " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'find the cart information
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListInvoiceAssetsbyCart " & request("idCart")
	set rsAssets = dbConnection.Execute(sql)
	
	'RRF - Set up test Variables to insert in to PDF Needs to be replaced with Database information
	TDate = Date()
	'RRF - Create the Object to the Tool Kit
	Set oTK = Server.CreateObject("APToolkit.Object")
	
	'Page 1 --------------------------------------------------------------------------------------------
	'Create The First Page
	
	'RRF - Select the first template PDF page
	sInputFile = "LAgreeP1.pdf"
	
	'RRF - Set the Output Filename
	r = oTK.OpenOutputFile(Server.MapPath(sUFileName))
	
	'RRF - Check for Errors for writing
	If r <> "0" Then
		Set oTK = Nothing
		response.write("Error: Cannot Open Output File, check NTFS write permissions AND IIS write permissions")
		response.end
	End If
	
	'RRF - Set Input file
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	
	'RRF - Check Read Errors
	If r <> "0" Then
		Set oTK = Nothing
		response.write("Error: Cannot Open Input File, check misspellings, NTFS read permissions AND IIS read permissions")
		response.end
	End If
	
	'RRF- Set the PDF Formatting is on
	oTK.DoFormFormatting = bDoFormFormatting
	
	'RRF - Set the Values - SetFormField (PDFFormField, Data, Attribute)
	r = oTK.SetFormFieldData ("Date", TDate, iFieldFlag)
	r = oTK.SetFormFieldData ("RPName", trim(rsCart("chrOSPerson")), iFieldFlag)
	r = oTK.SetFormFieldData ("RPEmail", "Need Email Field", iFieldFlag)
	r = oTK.SetFormFieldData ("RPPhone", trim(rsCart("chrOSPhone")), iFieldFlag)
	
	'RRF - Empty the rest of the fields, and copy to output page
	oTK.FlattenRemainingFormFields = True
	r = oTK.CopyForm(0, 0)
	
	if r < 1 Then
		Set oTK = Nothing
		response.write("Error: CopyForm Failed, possible bad input file, try doing a SaveAs in Acrobat")
		response.end
	end if
	
	'RRF - Reset the form, and create a new page on the output
	oTK.ResetFormFields
	oTK.NewPage
	
	'Start Page 2 ------------------------------------------------------------------------------------------
	'RRF - Select the 2nd Template Page
	sInputFile = "LAgreeP2.pdf"
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	oTK.DoFormFormatting = bDoFormFormatting
	
	r = oTK.SetFormFieldData ("OrderName", trim(rsCart("chrCart")), iFieldFlag)
	r = oTK.SetFormFieldData ("DeliverDate", formatdatetime(rsCart("dtShip"),2), iFieldFlag)
	r = oTK.SetFormFieldData ("ReturnDate", formatdatetime(rsCart("dtReturn"),2), iFieldFlag)
	
	oTK.FlattenRemainingFormFields = True
	r = oTK.CopyForm(0, 0)
	oTK.ResetFormFields
	oTK.NewPage
	
	'Start Page 3 -----------------------------------------------------------------------------------------
	'Select the 3rd Template Page.
	sInputFile = "LAgreeP3.pdf"
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	oTK.DoFormFormatting = bDoFormFormatting
	r = oTK.CopyForm(0, 0)
	oTK.NewPage
	
	'Start Page 4 -----------------------------------------------------------------------------------------
	'Select the 4th Template Page
	sInputFile = "4.pdf"
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	oTK.DoFormFormatting = bDoFormFormatting
	
	if rsAssets.EOF then
		oTK.FlattenRemainingFormFields = True
		r = oTK.CopyForm(0, 0)
		oTK.ResetFormFields
	else
		CurRec = 1
		do until rsAssets.EOF
			r = oTK.SetFormFieldData ("Qty" & CurRec, rsAssets("intOrdered"), FieldFlag)
			r = oTK.SetFormFieldData ("Item" & CurRec, trim(rsAssets("chrAssNum")), iFieldFlag)
			if rsAssets("chrType") = "C" then
				r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets("chrItem")) & ", " &_
				trim(rsAssets("chrProcessor"))& ", " & trim(rsAssets("chrMemory")) & ", " &_
				trim(rsAssets("chrHDD")) & ", " & trim(rsAssets("chrODrive")), iFieldFlag)
			else
				r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets("chrItem")), iFieldFlag)
			end if  
			r = oTK.SetFormFieldData ("Serial" & CurRec, trim(rsAssets("chrSerialNum")), iFieldFlag)
			if CurRec = 70 then
				oTK.FlattenRemainingFormFields = True
				r = oTK.CopyForm(0, 0)
				oTK.NewPage
				oTK.ResetFormFields
				CurRec = 1
				rsAssets.MoveNext
			else
				CurRec = CurRec + 1
				rsAssets.MoveNext
			end if
		loop
		oTK.FlattenRemainingFormFields = True
		r = oTK.CopyForm(0, 0)
		oTK.ResetFormFields
	end if
	oTK.CloseOutputFile
	
	'RRF - Email PDF ---------------------------------------------------------------------------------
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	Mailer.RemoteHost  = "63.236.44.26" 'mail server
	Mailer.FromName    = "administrator"
	Mailer.FromAddress = "administrator@techitsolutions.com"

	'AEG - Attach the User to the email
	Mailer.AddRecipient RPName, RPEmail
	Mailer.AddRecipient "Lesa", "lesa@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Start the Message information
	Mailer.Subject     = "Testing - PDF Loan Agreement"
	Mailer.BodyText    = "The loan agreement is attached." & VbCrLf & VbCrLf &_
	"Please fill out and return."

	mailer.AddAttachment Server.MapPath(sUFileName)

	'Execute the email
	Mailer.SendMail
	
	'Delete the Output file
	r = oTK.DeleteFile(Server.MapPath(sUFileName))
	Set oTK = Nothing
%>      