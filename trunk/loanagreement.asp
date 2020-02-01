<%@ Language=VBScript %>
<!-- #include file="includes/openconn.asp" -->
<%
'RRF - Set Variables for Active Tool Kit
'RRF - MaxRowsPerPage is the number of Fields on the 4th Page the inventory page of the PDF
	MaxRowsPerPage = 70
'RRF - iFieldFlag is a attribute code that make the field flatten and keep the imported data.
	iFieldFlag = -998
'RRF - bDoFormFormatting will keep the text formats that is designated on the PDF fields
	bDoFormFormatting = False
'RRF - sUFileName is the Name of the Final PDF file to send to the customer.  
	'sUFileName = "test.pdf"
	sUFileName = "LoanAgreement-" & request("idCart") & ".pdf"
    
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

'RRF - Set the Output Filename
	r = oTK.OpenOutputFile(Server.MapPath(sUFileName))

'RRF - Check for Errors for writing
	If r <> "0" Then
		Set oTK = Nothing
		response.write("Error: Cannot Open Output File, check NTFS write permissions AND IIS write permissions")
		response.end
	End If
	
'RRF - Check Read Errors
	If r <> "0" Then
		Set oTK = Nothing
		response.write("Error: Cannot Open Input File, check misspellings, NTFS read permissions AND IIS read permissions")
		response.end
	End If
  
'RRF - Set Input file
'AEG - check to see which Customer it is
	select case rsCart("idCustomer")
		case 1 'Corp Events
			sInputFile = "LAgreeP1.pdf"
		case 2 'Apple Edu
			sInputFile = "LAgreeP1e.pdf"
		case 3 'techIT
			sInputFile = "LAgreeP1.pdf"
		case 6 'NED
			sInputFile = "LAgreeP1e.pdf"
		case 7 'SED
			sInputFile = "LAgreeP1e.pdf"
		case 8 'Macromedia
			sInputFile = "LAgreeP1.pdf"
		case 10 'IDG
			sInputFile = "LAgreeP1.pdf"
		case 11 'XED
			sInputFile = "LAgreeP1e.pdf"
		case 13 'APS
			sInputFile = "LAgreeP1.pdf"
		case 15 'EPS
			sInputFile = "LAgreeP1.pdf"
		case else
			sInputFile = "LAgreeP1.pdf"
	end select
	'sInputFile = "LAgreeP1.pdf"
'AEG - Open the fist file
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
'RRF- Set the PDF Formatting is on
	oTK.DoFormFormatting = bDoFormFormatting

'RRF - Set the Values - SetFormField (PDFFormField, Data, Attribute)
	r = oTK.SetFormFieldData ("Date", TDate, iFieldFlag)
	r = oTK.SetFormFieldData ("RPName", trim(rsCart("chrOSPerson")), iFieldFlag)
	r = oTK.SetFormFieldData ("RPEmail", trim(rsCart("chrOSEmail")), iFieldFlag)
	r = oTK.SetFormFieldData ("RPPhone", trim(rsCart("chrOSPhone")), iFieldFlag)

'RRF - Empty the rest of the fields, and copy to output page
	oTK.FlattenRemainingFormFields = False
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
'AEG - check to see which Customer it is
	select case rsCart("idCustomer")
		case 1 'Corp Events
			sInputFile = "LAgreeP2.pdf"
		case 2 'Apple Edu
			sInputFile = "LAgreeP2e.pdf"
		case 3 'techIT
			sInputFile = "LAgreeP2.pdf"
		case 6 'NED
			sInputFile = "LAgreeP2e.pdf"
		case 7 'SED
			sInputFile = "LAgreeP2e.pdf"
		case 8 'Macromedia
			sInputFile = "LAgreeP2.pdf"
		case 10 'IDG
			sInputFile = "LAgreeP2.pdf"
		case 11 'XED
			sInputFile = "LAgreeP2e.pdf"
		case 13 'APS
			sInputFile = "LAgreeP2.pdf"
		case 15 'EPS
			sInputFile = "LAgreeP2.pdf"
		case else
			sInputFile = "LAgreeP2.pdf"
	end select
	'sInputFile = "LAgreeP2.pdf"
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	oTK.DoFormFormatting = bDoFormFormatting

	r = oTK.SetFormFieldData ("RPName", trim(rsCart("chrOSPerson")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress", trim(rsCart("chrAddress")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress2", trim(rsCart("chrAddress2")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress3", trim(rsCart("chrAddress3")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress4", trim(rsCart("chrAddress4")), iFieldFlag)
	chrCityStateZip = trim(rsCart("chrCity")) & ", " & trim(rsCart("chrState")) & " " & trim(rsCart("chrZip"))
	r = oTK.SetFormFieldData ("chrCityStateZip", chrCityStateZip, iFieldFlag)
	
	r = oTK.SetFormFieldData ("chrOSPerson", trim(rsCart("chrOSPerson")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrOSEmail", trim(rsCart("chrOSEmail")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrOSPhone", trim(rsCart("chrOSPhone")), iFieldFlag)
	
	r = oTK.SetFormFieldData ("OrderName", trim(rsCart("chrCart")), iFieldFlag)
	r = oTK.SetFormFieldData ("DeliverDate", formatdatetime(rsCart("dtArrival"),2), iFieldFlag)
	r = oTK.SetFormFieldData ("ReturnDate", formatdatetime(rsCart("dtReturn"),2), iFieldFlag)
	
	oTK.FlattenRemainingFormFields = True
	r = oTK.CopyForm(0, 0)
	
	if r < 1 Then
		Set oTK = Nothing
		response.write("Error: CopyForm Failed, possible bad input file, try doing a SaveAs in Acrobat")
		response.end
	end if
	
	oTK.ResetFormFields
	oTK.NewPage

'Start Page 3 -----------------------------------------------------------------------------------------
'Select the 3rd Template Page.
'AEG - check to see which Customer it is
	select case rsCart("idCustomer")
		case 1 'Corp Events
			sInputFile = "LAgreeP3.pdf"
		case 2 'Apple Edu
			sInputFile = "LAgreeP3e.pdf"
		case 3 'techIT
			sInputFile = "LAgreeP3.pdf"
		case 6 'NED
			sInputFile = "LAgreeP3e.pdf"
		case 7 'SED
			sInputFile = "LAgreeP3e.pdf"
		case 8 'Macromedia
			sInputFile = "LAgreeP3.pdf"
		case 10 'IDG
			sInputFile = "LAgreeP3.pdf"
		case 11 'XED
			sInputFile = "LAgreeP3e.pdf"
		case 13 'APS
			sInputFile = "LAgreeP3.pdf"
		case 15 'EPS
			sInputFile = "LAgreeP3.pdf"
		case else
			sInputFile = "LAgreeP3.pdf"
	end select
	'sInputFile = "LAgreeP3.pdf"
	r = oTK.OpenInputFile(Server.MapPath(sInputFile))
	oTK.DoFormFormatting = bDoFormFormatting
	
	r = oTK.SetFormFieldData ("chrAddress", trim(rsCart("chrAddress")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress2", trim(rsCart("chrAddress2")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress3", trim(rsCart("chrAddress3")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrAddress4", trim(rsCart("chrAddress4")), iFieldFlag)
	chrCityStateZip = trim(rsCart("chrCity")) & ", " & trim(rsCart("chrState")) & " " & trim(rsCart("chrZip"))
	r = oTK.SetFormFieldData ("chrCityStateZip", chrCityStateZip, iFieldFlag)
	
	r = oTK.SetFormFieldData ("chrOSPerson", trim(rsCart("chrOSPerson")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrOSEmail", trim(rsCart("chrOSEmail")), iFieldFlag)
	r = oTK.SetFormFieldData ("chrOSPhone", trim(rsCart("chrOSPhone")), iFieldFlag)
	
	r = oTK.CopyForm(0, 0)
	oTK.NewPage

'Start Page 4 -----------------------------------------------------------------------------------------
'Select the 4th Template Page
'AEG - check to see which Customer it is
	select case rsCart("idCustomer")
		case 1 'Corp Events
			sInputFile = "LAgreeP4.pdf"
		case 2 'Apple Edu
			sInputFile = "LAgreeP4e.pdf"
		case 3 'techIT
			sInputFile = "LAgreeP4.pdf"
		case 6 'NED
			sInputFile = "LAgreeP4e.pdf"
		case 7 'SED
			sInputFile = "LAgreeP4e.pdf"
		case 8 'Macromedia
			sInputFile = "LAgreeP4.pdf"
		case 10 'IDG
			sInputFile = "LAgreeP4.pdf"
		case 11 'XED
			sInputFile = "LAgreeP4e.pdf"
		case 13 'APS
			sInputFile = "LAgreeP4.pdf"
		case 15 'EPS
			sInputFile = "LAgreeP4.pdf"
		case else
			sInputFile = "LAgreeP4.pdf"
	end select
	'sInputFile = "LAgreeP4.pdf"
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
				r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets("chrItem")) & ", " & trim(rsAssets("chrProcessor"))& ", " & trim(rsAssets("chrMemory")) & ", " & trim(rsAssets("chrHDD")) & ", " & trim(rsAssets("chrODrive")), iFieldFlag)
			else
				r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets("chrItem")), iFieldFlag)
			end if
			r = oTK.SetFormFieldData ("Serial" & CurRec, trim(rsAssets("chrSerialNum")), iFieldFlag)
			'this is the maximum number of Fields on the PDF
			if CurRec = 43 then
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
%>
	<SCRIPT LANGUAGE="JavaScript">
		document.location.href="<%response.write(sUFileName)%>"
	</SCRIPT>
<%        
'Delete the Output file
	'r = oTK.DeleteFile(Server.MapPath(sUFileName))
	Set oTK = Nothing
%>      