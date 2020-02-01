<%
' name of the input template to use
  sInputFile = "LAgreeP4.pdf"

' connection string to the MS Access DB
 ' sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("northwind.mdb") & ";Persist Security Info=False"

' the number of rows of form fields we created on our input tempate
  iRowsPerPage = 30

' Number of records to print
  rnum = 200

' the field flag. (determines whether fields are to be removed, read-only, editable, or hidden)
  iFieldFlag = -998

' the form formatting. (determines whether pre-defined formatting is retained in the outputted pdf)
  bDoFormFormatting = True

' build the SQL Statement from the previous form
  ' If sCountry <> "*" Then
  '    sSQL = "Select ContactName, CompanyName, ContactTitle, City, Country, Phone From Customers where Country = '" & sCountry & "';"
  ' Else
  '    sSQL = "Select ContactName, CompanyName, ContactTitle, City, Country, Phone From Customers"
  ' End If

' create our Toolkit Object
  Set oTK = Server.CreateObject("APToolkit.Object")
  
  NumPage = 1
  Rrec = 1
  
  Do until Rrec >= 200
'Create a Unique File Name and open the output file
    sUFileName = NumPage & ".pdf"
    r = oTK.OpenOutputFile(Server.MapPath(sUFileName))
    If r <> "0" Then
      Set oTK = Nothing
      response.write("Error: Cannot Open Output File, check NTFS write permissions AND IIS write permissions")
      response.end
    End If

' Open the existing input file that we have created
    r = oTK.OpenInputFile(Server.MapPath(sInputFile))
    If r <> "0" Then
      Set oTK = Nothing
      response.write("Error: Cannot Open Input File, check misspellings, NTFS read permissions AND IIS read permissions")
      response.end
    End If

' let us open our connection to the database, pass the query, and receive our recordset
 ' Set oDBConn = CreateObject("ADODB.Connection")
 ' oDBConn.Open sConn
 ' Set oRS = oDBConn.Execute(sSQL)

' loop through the recordset until there are no more records
  For i = 1 to iRowsPerPage 'predefined
    oTK.DoFormFormatting = bDoFormFormatting
    r = oTK.SetFormFieldData ("Qty" & i, Rrec, iFieldFlag)
    Rrec = Rrec + 1
  Next

' we will now flatten any of the form fields that did not receive values
  oTK.FlattenRemainingFormFields = True

' here is where we will implement all the changes and actually create the PDF
  r = oTK.CopyForm(0, 0)
  if r < 1 Then
    Set oTK = Nothing
    response.write("Error: CopyForm Failed, possible bad input file, try doing a SaveAs in Acrobat")
    response.end
  end if

' we reset the form fields so the values will not be the same
  oTK.ResetFormFields

' close the output file, we are finished
  oTK.CloseOutputFile
  NumPage = NumPage + 1
  Loop

  Set oTK = Nothing
' below we will use most failsafe method of redirecting a PDF to the browser
%>
        <SCRIPT LANGUAGE="JavaScript">
        document.location.href="<%response.write(sUFileName)%>"
        </SCRIPT>