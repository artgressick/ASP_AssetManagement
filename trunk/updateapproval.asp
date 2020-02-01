<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	idStatus = 3 'this is approved
	
	'Check to see if the check box was checked we will send an email at the end.
	if request("idLoaner") = 1 then
		idLoaner = 1
	else
		idLoaner = 0
	end if
	
	'Send the information string with contents from the form.
	sql = "execute UpdateApprovalBilling2 " & _
		request("idCart") & "," &_
		idStatus & "," &_
		request("idBillto") & "," &_
		idLoaner & "," &_
		request("idExpedite") & "," &_
		session("idUser")
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'AEG - Find the Cart Name so that we can send an Email
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	idCart = rsCart("idCart")
	chrCart = trim(rsCart("chrCart"))
	idCustomer = rsCart("idCustomer")
	
	'AEG - Close the Recordset
	rsCart.Close
	set rsCart = nothing
	
	'--------------------------------------------------------------------------
	'AEG - Open the SMTP Mailer Client
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	Mailer.RemoteHost  = "techit-ex2.techitsolutions.com" 'mail server
	Mailer.FromName    = "administrator"
	Mailer.FromAddress = "administrator@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Find the User to send an email.
	set rsUser = server.CreateObject("adodb.recordset")
	sql = "execute FindUserWhoEnteredCartbyCartID " & request("idCart")
	set rsUser = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrUserName = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
	chrUserEmail = trim(rsUser("chrEmail"))
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrUserName, chrUserEmail
	
	'Close the User Recordset
	rsUser.Close
	set rsUser = nothing
	'--------------------------------------------------------------------------
	'AEG - Find the Warehouse to send an email.
	set rsWarehouse = server.CreateObject("adodb.recordset")
	sql = "execute FindWarehouseEmailbyCartID " & request("idCart")
	set rsWarehouse = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrWarehouseName = trim(rsWarehouse("chrEmailName"))
	chrWarehouseEmail = trim(rsWarehouse("chrEmailAddress"))
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrWarehouseName, chrWarehouseEmail
	
	'Close the User Recordset
	rsWarehouse.Close
	set rsWarehouse = nothing
	'--------------------------------------------------------------------------
	'AEG - Start the Message information
	Mailer.Subject     = "Cart Approved - " & chrCart
	Mailer.BodyText    = chrUserName & "," & VbCrLf & VbCrLf &_
	"Your cart, " & chrCart & " has been approved by the Pool Manager." & VbCrLf & VbCrLf &_
	"Please contact your Pool Manager for any additional information regarding this cart." & VbCrLf & VbCrLf &_
	"Thank you..." & VbCrLf &_
	"techIT Solutions Asset Management Team"
	
	'Execute the email
	Mailer.SendMail
	'clear out the mailer
	mailer.ClearAllRecipients
	mailer.ClearBodyText
	
	'Begin the email for the Loaner agreement if needed
	if idLoaner = 1 then
		'RRF - Set Variables for Active Tool Kit
		'RRF - MaxRowsPerPage is the number of Fields on the 4th Page the inventory page of the PDF
		MaxRowsPerPage = 70
		'RRF - iFieldFlag is a attribute code that make the field flatten and keep the imported data.
		iFieldFlag = -998
		'RRF - bDoFormFormatting will keep the text formats that is designated on the PDF fields
		bDoFormFormatting = False
		'RRF - sUFileName is the Name of the Final PDF file to send to the customer.  
		sUFileName = "LoanAgreement-" & idCart & ".pdf"
		
		'find the cart information
		set rsCart2 = server.CreateObject("adodb.recordset")
		sql = "execute ViewInvoicebyCart " & idCart
		set rsCart2 = dbConnection.Execute(sql)
		
		'find the cart information
		set rsAssets2 = server.CreateObject("adodb.recordset")
		sql = "execute ListInvoiceAssetsbyCart " & idCart
		set rsAssets2 = dbConnection.Execute(sql)
		
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
		
		'RRF - Select the first template PDF page
		'AEG - check to see which Customer it is
		select case idCustomer
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
		
		'RRF - Set Input file
		r = oTK.OpenInputFile(Server.MapPath(sInputFile))
		
		'RRF- Set the PDF Formatting is on
		oTK.DoFormFormatting = bDoFormFormatting
		
		'RRF - Set the Values - SetFormField (PDFFormField, Data, Attribute)
		r = oTK.SetFormFieldData ("Date", TDate, iFieldFlag)
		r = oTK.SetFormFieldData ("RPName", trim(request("chrOSPerson")), iFieldFlag)
		r = oTK.SetFormFieldData ("RPEmail", trim(request("chrOSEmail")), iFieldFlag)
		r = oTK.SetFormFieldData ("RPPhone", trim(rsCart2("chrOSPhone")), iFieldFlag)
		
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
		'AEG - check to see which Customer it is
		select case idCustomer
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
		
		r = oTK.SetFormFieldData ("RPName", trim(rsCart2("chrOSPerson")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress", trim(rsCart2("chrAddress")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress2", trim(rsCart2("chrAddress2")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress3", trim(rsCart2("chrAddress3")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress4", trim(rsCart2("chrAddress4")), iFieldFlag)
		chrCityStateZip = trim(rsCart2("chrCity")) & ", " & trim(rsCart2("chrState")) & " " & trim(rsCart2("chrZip"))
		r = oTK.SetFormFieldData ("chrCityStateZip", chrCityStateZip, iFieldFlag)
		
		r = oTK.SetFormFieldData ("chrOSPerson", trim(rsCart2("chrOSPerson")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrOSEmail", trim(rsCart2("chrOSEmail")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrOSPhone", trim(rsCart2("chrOSPhone")), iFieldFlag)
		
		r = oTK.SetFormFieldData ("OrderName", trim(rsCart2("chrCart")), iFieldFlag)
		r = oTK.SetFormFieldData ("DeliverDate", formatdatetime(rsCart2("dtArrival"),2), iFieldFlag)
		r = oTK.SetFormFieldData ("ReturnDate", formatdatetime(rsCart2("dtReturn"),2), iFieldFlag)
		
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
		select case idCustomer
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
		
		r = oTK.SetFormFieldData ("chrAddress", trim(rsCart2("chrAddress")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress2", trim(rsCart2("chrAddress2")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress3", trim(rsCart2("chrAddress3")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrAddress4", trim(rsCart2("chrAddress4")), iFieldFlag)
		chrCityStateZip = trim(rsCart2("chrCity")) & ", " & trim(rsCart2("chrState")) & " " & trim(rsCart2("chrZip"))
		r = oTK.SetFormFieldData ("chrCityStateZip", chrCityStateZip, iFieldFlag)
		
		r = oTK.SetFormFieldData ("chrOSPerson", trim(rsCart2("chrOSPerson")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrOSEmail", trim(rsCart2("chrOSEmail")), iFieldFlag)
		r = oTK.SetFormFieldData ("chrOSPhone", trim(rsCart2("chrOSPhone")), iFieldFlag)
		
		r = oTK.CopyForm(0, 0)
		oTK.NewPage
		
		'Start Page 4 -----------------------------------------------------------------------------------------
		'Select the 4th Template Page
		'AEG - check to see which Customer it is
		select case idCustomer
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
		
		if rsAssets2.EOF then
			oTK.FlattenRemainingFormFields = True
			r = oTK.CopyForm(0, 0)
			oTK.ResetFormFields
		else
			CurRec = 1
			do until rsAssets2.EOF
				r = oTK.SetFormFieldData ("Qty" & CurRec, rsAssets2("intOrdered"), FieldFlag)
				r = oTK.SetFormFieldData ("Item" & CurRec, trim(rsAssets2("chrAssNum")), iFieldFlag)
				if rsAssets2("chrType") = "C" then
					r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets2("chrItem")) & ", " & trim(rsAssets2("chrProcessor"))& ", " & trim(rsAssets2("chrMemory")) & ", " & trim(rsAssets2("chrHDD")) & ", " & trim(rsAssets2("chrODrive")), iFieldFlag)
				else
					r = oTK.SetFormFieldData ("Info" & CurRec, trim(rsAssets2("chrItem")), iFieldFlag)
				end if
				r = oTK.SetFormFieldData ("Serial" & CurRec, trim(rsAssets2("chrSerialNum")), iFieldFlag)
				if CurRec = 43 then
					oTK.FlattenRemainingFormFields = True
					r = oTK.CopyForm(0, 0)
					oTK.NewPage
					oTK.ResetFormFields
					CurRec = 1
					rsAssets2.MoveNext
				else
					CurRec = CurRec + 1
					rsAssets2.MoveNext
				end if
			loop
			oTK.FlattenRemainingFormFields = True
			r = oTK.CopyForm(0, 0)
			oTK.ResetFormFields
		end if
		oTK.CloseOutputFile
		
		'AEG - Attach the User to the email
		Mailer.AddRecipient trim(request("chrOSPerson")), trim(request("chrOSEmail"))
		Mailer.AddRecipient "techIT Operations", "operations@techitsolutions.com"
		'AEG - check to see which Customer it is
		select case idCustomer
			case 1 'Corp Events
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
			case 2 'Apple Edu
				Mailer.AddRecipient "David Hitchcock", "dhitchcock@techitsolutions.com"
				Mailer.AddRecipient "William Hernderson", "whenderson@techitsolutions.com"
			case 3 'techIT
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
			case 6 'NED
				Mailer.AddRecipient "David Hitchcock", "dhitchcock@techitsolutions.com"
				Mailer.AddRecipient "William Hernderson", "whenderson@techitsolutions.com"
			case 7 'SED
				Mailer.AddRecipient "David Hitchcock", "dhitchcock@techitsolutions.com"
				Mailer.AddRecipient "William Hernderson", "whenderson@techitsolutions.com"
			case 8 'Macromedia
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
			case 10 'IDG
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
			case 11 'XED
				Mailer.AddRecipient "David Hitchcock", "dhitchcock@techitsolutions.com"
				Mailer.AddRecipient "William Hernderson", "whenderson@techitsolutions.com"
			case 13 'APS
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
			case 15 'EPS
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
			case else
				Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
				Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
				Mailer.AddRecipient "Chris Angerame", "cangerame@techitsolutions.com"
		end select
		'Mailer.AddRecipient "Dave Arnold", "darnold@techitsolutions.com"
		'Mailer.AddRecipient "Tim Reed", "treed@techitsolutions.com"
		'--------------------------------------------------------------------------
		'AEG - Start the Message information
		Mailer.Subject     = "Loan Agreement for " & chrCart
		Mailer.BodyText    = "Attached is the Equipment Loan Agreement form. Please fill out and return as soon as possible. Your equipment will not ship until this form has been returned."
		
		mailer.AddAttachment Server.MapPath(sUFileName)
		
		'Execute the email
		Mailer.SendMail
		
		'Delete the Output file
		r = oTK.DeleteFile(Server.MapPath(sUFileName))
		Set oTK = Nothing
		
		rsCart2.Close
		set rsCart2 = nothing
		rsAssets2.close
		set rsAssets2 = nothing
	end if
	
	'Close database connections
	
	dbConnection.Close
	set dbConnection = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Cart Approved !!</strong></font></td>
                </tr>
                <tr> 
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">This cart has been approved and sent to our Warehouse Team for processing.<br><br>
				  Please click <a href="orders.asp">here</a> to return to the Orders Section.</font></td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright &copy; 2003 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.1 &amp; Corporate Business 
                    Intelligence are products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
  </table>
</body>
</html>