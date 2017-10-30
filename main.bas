Attribute VB_Name = "Module1"


Dim OutcomeVal As String
Dim TotalRows As Long
Dim RowNumber As Long
Dim i As Long
Dim j As Long
Dim headerValues() As Variant
Dim sht As Excel.Worksheet
Dim rowCount As Long
Dim colCount As Long
'Vars for MsgBox
Dim WeekEnding As String
Dim iRet As Integer
Dim strPrompt As String
Dim strTitle As String
'CurrentDate to check expiry date of credit cards
Dim CurrentDate As Date
'To generate Batch Month for reporting purposes
Dim strTemp As String
Dim index As String
Dim finalString As String
Dim Test1 As Variant
Dim Temp1 As Variant
Dim Temp2 As Variant
Dim campaignName As String
Dim wb As Workbook
Dim current As Excel.Worksheet
Dim oldColIndex As Long
Dim newColIndex As Long
Dim colMap() As String
Dim Initiative_Name As String
Dim Initiative_Type As String

'Global Cells Defined globally to allow for simple methods
'-------------------------------------------------------
Dim Current_Cell As Object
Dim Gender_Cell As Object
Dim Title_Cell As Object
Dim Outcome_Cell As Object
Dim Payment_Method_Cell As Object
Dim Card_Type_Cell As Object
Dim Card_Number_Cell As Object
Dim Card_Expiry_Cell As Object
Dim Card_Name_Cell As Object
Dim BSB_Cell As Object
Dim Bank_Name_Cell As Object
Dim Account_Num_Cell As Object
Dim Account_Name_Cell As Object
'-------------------------------------------------------
'Simple boolean marked true if there is an error in the macro
Dim Error_Has_Occurred As Boolean
Dim colName As String
Dim Charity As String
Dim ListMatch As Object
Dim Column_Not_Found As String
Dim DataRange As Variant
'PETER MAC ID
Dim PETER_MAC_FILE_ID As Integer


'|--------------------------------------------------------------------------------------------------------------------|
'|---------------------------------------DATA CLEANING CONTEXT PREPERATION--------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

Sub WAP_Generate_AONSW()
    'HOUSEKEEPING
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    'Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    'Set wb = Workbooks.Open("D:\Python\api\headers.csv")
    'delete data from sheet WAP
    PETER_MAC_FILE_ID = Workbooks("WAYS Data Cleaning Macro").Sheets("PM_ID_NUM").Cells(1, 1)
    Workbooks("WAYS Data Cleaning Macro").Sheets("DATA").Activate
    Workbooks("WAYS Data Cleaning Macro").Sheets("DATA").Range("A:AAA").Delete
    Initiative_Name = Workbooks("WAYS Data Cleaning Macro").Sheets("Macro").ComboBox1.Value
    nameSplit = Split(Initiative_Name, " ")
    Charity = nameSplit(0)
    Initiative_Type = nameSplit(1)
    For i = 2 To UBound(nameSplit)
        Initiative_Type = Initiative_Type + " " + Split(Initiative_Name, " ")(i)
    Next i
    Set ListMatch = CreateObject("vbscript.regexp")
    ListMatch.Pattern = "[0-9]{6}"
    
    Sheets("Macro").Activate
    Error_Has_Occurred = False
    'Find csv file to import
    strFilename = Application.GetOpenFilename
    Workbooks("WAYS Data Cleaning Macro").Sheets("DATA").Activate
        
    'apply text formatting data
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & strFilename & "" _
        , Destination:=Range("$A$1"))
        .name = ""
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
        
        On Error Resume Next
        
        'Copy to new workbook
        Worksheets("DATA").Cells.Select
        Selection.Copy
        Workbooks.Add
        ActiveSheet.Paste
        ActiveSheet.name = "Sheet1"
        
        'Delete Blank Rows first
        ActiveSheet.UsedRange.Select
        With Application
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            For i = Selection.Rows.Count To 2 Step -1
                If WorksheetFunction.CountA(Selection.Rows(i)) = 0 Then
                    Selection.Rows(i).EntireRow.Delete
                End If
            Next i
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
        End With
        
        'Generates WAYS Version"

        Sheets("Sheet1").Activate
        Generate_Column_Map
        Write_Headers
        CopyPaste
        DataCleaning
        If Initiative_Name = "STC Inbound" Then
            Generate_STC_Core_Data
        Else
            Generate_Core_Data
        End If
        Choose_Data_Generation
        Highlight_Headers
        Sheets("Sheet1").Delete
        Sheets(Initiative_Name).Cells.EntireColumn.AutoFit
        Sheets(Initiative_Name).Range("A1").AutoFilter
    
    If (Error_Has_Occurred = False) Then
        'SHOW COMPLETED MESSAGEBOX
        strPrompt = Initiative_Name + " (WAYS Version): IT HAS BEEN GENERATED!!"
        strTitle = "Success"
        iRet = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
    End If
        
    
    'HOUSEKEEPING
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    'Application.Calculation = calcState
    Application.EnableEvents = True

End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|---------------------------------------------END OF DATA PREPERATION------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-----------------------------------------------COLUMN MAPPING-------------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO UTILIZES THE HIDDEN SHEETS LABELLED WITH THE CHARITY NAME TO CREAT A MAPPING OF       |
'|    ORIGINAL COLUMN : NEW COLUMN. THIS MAPPING ALLOWS FOR DYNAMIC GENERATION OF COLUMNS AND ENSURES REARRANGING     |
'|    IS TRIVIAL. IN ORDER FOR THE COLUMN MAPPING TO BE SUCCESSFUL, YOU MUST ENSURE THAT THE NAMES IN THE ORIGINAL    |
'|    COLUMN ORDERING IS EXACTLY THE SAME (INCLUDING CASE) AS THE NEW COLUMN ORDERING.                                |
'|____________________________________________________________________________________________________________________|

Sub Generate_Column_Map()
    Dim counter As Integer
    newTotal = Row_Length(1)
    oldTotal = Row_Length(2)
    mapping = ""
    ReDim colMap(newTotal)
    For l = 1 To newTotal
        counter = 0
        For k = 1 To oldTotal
           If Workbooks("WAYS Data Cleaning Macro").Sheets(Initiative_Name).Cells(1, l) = Workbooks("WAYS Data Cleaning Macro").Sheets(Initiative_Name).Cells(2, k) Then
            mapping = CStr(l) + " " + CStr(k)
            colMap(l - 1) = mapping
            counter = counter + 1
           End If
        Next k
        If counter = 0 Then
            colMap(l - 1) = CStr(l) + " " + "512"
        End If
    Next l
End Sub

Sub CopyPaste()
    Sheets(1).Activate
    TotalColumns = Row_Length(1) - 1
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    For P = 0 To TotalColumns
        oldColPos = CInt(Split(colMap(P), " ")(1))
        newColPos = CInt(Split(colMap(P), " ")(0))
        For RowNumber = 2 To TotalRows
            Sheets(Initiative_Name).Cells(RowNumber, newColPos) = Cells(RowNumber, oldColPos)
        Next RowNumber
    Next P
End Sub

Public Function Row_Length(rowNum As Integer)
    Row_Length = 1
    EndOfRow = False
    While (EndOfRow = False)
        If Workbooks("WAYS Data Cleaning Macro").Sheets(Initiative_Name).Cells(rowNum, Row_Length) <> "" Then
             Row_Length = Row_Length + 1
        Else
            EndOfRow = True
        End If
    Wend
End Function

Sub Write_Headers()
    Sheets.Add After:=ActiveSheet
    Sheets(2).Select
    Sheets(2).name = Initiative_Name
    Set current = Worksheets(Initiative_Name)
    Totalcol = Row_Length(1)
    For ColNumber = Totalcol To 1 Step -1
        Worksheets(Initiative_Name).Cells(1, ColNumber) = Workbooks("WAYS Data Cleaning Macro").Sheets(Initiative_Name).Cells(1, ColNumber)
    Next ColNumber
    PopulateHeaderValues
End Sub

'sets Global Variable headerValues
Sub PopulateHeaderValues()
 headerMax = current.UsedRange.Columns.Count
 ReDim headerValues(headerMax)
 For i = 0 To headerMax Step 1
    headerValues(i) = current.Cells(1, (i + 1))
 Next i
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|---------------------------------------------END OF COLUMN MAPPING--------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|------------------------------------------------DATA CLEANING-------------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO FOCUSES ON CLEANING SINGLE COLUMNS OF DATA. IF, IN ORDER TO CLEAN THE COLUMN,         |
'|     YOU NEED TO KNOW THE VALUE IN ANOTHER COLUMN, THEN THAT WILL BE PLACED IN THE DATA GENERATION SECTION.         |
'|____________________________________________________________________________________________________________________|

'This is the main parent method to do basic single column cleaning. The macro determines how to clean a column by
'looking at the 3rd row in the associated sheet, to determine what type the column is
Sub PMac_Id_Box()
    Dim ID As String
    ID = InputBox("Please Enter the starting ID for this file", "Peter Mac ID Creation", CStr(PETER_MAC_FILE_ID))
    PETER_MAC_FILE_ID = Int(ID)

End Sub

Sub DataCleaning()
    On Error GoTo HandleError
    Dim colNum As Integer
    Dim colType As String
    If Charity = "PM" Then
        PMac_Id_Box
    End If
    TotalRows = Sheets(1).UsedRange.Rows.Count
    CurrentDate = Date
    colCount = Worksheets(Initiative_Name).UsedRange.Columns.Count
    For C = 1 To colCount - 1
        'colType refers to the text found in the 3rd row of associated sheet. if colType = "Address" the macro will
        'jump to the address if statement shown below.
        colType = Workbooks("WAYS Data Cleaning Macro").Sheets(Initiative_Name).Cells(3, C)
        colName = Worksheets(Initiative_Name).Cells(1, C)
        colNum = C
        For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        
            'this is an important feature of the single column cleaning. Current_Cell is set here as a global variable
            'meaning that it can be accessed from anywhere in the macro. Therefore whenever you see current cell in this
            'section or the 'General Functions' section this is where the value is set.
            Set Current_Cell = Worksheets(Initiative_Name).Cells(currentRow, colNum)
            
            'Not Null is used as a type for any column which should never have blank cells.
            If (colType = "Not Null") Then
                Clean_Not_Null
            
            ElseIf colType = "Title" Then
                Clean_Title
                
            ElseIf colType = "Name" Then
                Clean_Name
            ElseIf colType = "Payment Name" Then
                Clean_Payment_Name
                
            ElseIf colType = "DOB" Then
                Clean_DOB
                
            ElseIf colType = "Gender" Then
                Clean_Gender
                
            ElseIf colType = "Address1" Then
                Clean_Address1
                
            ElseIf colType = "Suburb" Then
                Clean_Suburb
                
            ElseIf colType = "State" Then
                Clean_State
                
            ElseIf colType = "Postcode" Then
                Clean_Postcode
                
            ElseIf colType = "Phone" Then
                Clean_Phone
                
            ElseIf colType = "RG" Then
                Clean_Regular_Gift (currentRow)
                
            ElseIf colType = "SG" Then
                Clean_Single_Gift (currentRow)
                
            ElseIf colType = "Start Month" Then
                Clean_Start_Month (currentRow)
                
            ElseIf colType = "Campaign" Then
                Clean_Campaign
                
            ElseIf colType = "cc num" Then
                Clean_Credit_Card_Number
                
            ElseIf colType = "cc expiry" Then
                Clean_Expiry_Date
                
            ElseIf colType = "dd num" Then
                Clean_DD_Num
            'amnesty specific
            ElseIf colType = "aia campaign" Then
                Clean_AIA_Campaign
                
            'amnesty specific
            ElseIf colType = "aia program code" Then
                Clean_AIA_Program_Code
            'amnesty specific
            ElseIf colType = "aia how taken" Then
                Clean_AIA_How_Taken
            
            'amnesty specific
            ElseIf colType = "aia source" Then
                Clean_AIA_CC_Source
                
            'amnesty specific
            ElseIf colType = "aia audience" Then
                Clean_AIA_Audience
                
            'amnesty specific
            ElseIf colType = "aia action type" Then
                Clean_AIA_Action_Type
                
            ElseIf colType = "Frequency" Then
                Clean_Frequency (currentRow)
                
            ElseIf colType = "Debit Day" Then
                Clean_Debit_Day (currentRow)
                
            ElseIf colType = "Outcome" Then
                Clean_Outcomes
                
            ElseIf colType = "WWF_Lead_New_Details" Then
                Clean_WWF_New_Details
            ElseIf colType = "Date" Then
                Clean_Date
                
            ElseIf colType = "PMac ID" Then
                Copy_ID_Pmac (currentRow)
            ElseIf colType = "Number" Then
                Convert_To_Number
                
            End If
        Next currentRow

    Next C
Exit Sub
'This directs the macro to the procedrue "Handle_Error" found at the bottom of the macro in the
'General Functions section.
HandleError:
    Handle_Error ("Data Cleaning")
    Exit Sub
End Sub
Sub Convert_To_Number()
    If Len(Trim(Current_Cell)) > 0 Then
        Current_Cell = Format(Current_Cell, "##.00")
        
    End If

End Sub
Sub Clean_DD_Num()
    If (Len(Trim(Current_Cell)) > 9) Then
        Highlight_Yellow
    End If
End Sub

Sub Copy_ID_Pmac(currentRow As Integer)
    Dim ID_Cell As Object
    Set ID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1")))
    
    Current_Cell = ID_Cell
End Sub

Sub Clean_Date()
    Current_Cell.NumberFormat = "@"
    Current_Cell = Replace(Current_Cell, ".", "/")
    Current_Cell = Format(Current_Cell, "dd/mm/yyyy")
    
End Sub

Sub Clean_WWF_New_Details()
    If Len(Trim(Current_Cell)) <> 0 Then
        Highlight_Yellow
    End If
End Sub

Sub Clean_AIA_CC_Source()
    Current_Cell = "2016 Donor Conversion"
End Sub

Sub Clean_AIA_Action_Type()
    Current_Cell = "Petition"
End Sub

Sub Clean_AIA_Audience()
    Current_Cell = "INDIV"
End Sub

Sub Clean_AIA_How_Taken()
    Current_Cell = "TM"
End Sub

Sub Clean_AIA_Program_Code()
    If Initiative_Type <> "Cash Conversion" Then
        Current_Cell = "556"
    Else
        Current_Cell = "725"
    End If
End Sub

Sub Clean_AIA_Campaign()
    If Initiative_Type <> "Cash Conversion" Then
        Current_Cell = "W4R"
    Else
        Current_Cell = "IAR"
    End If
End Sub

Sub Clean_Campaign()
    Current_Cell = Initiative_Name
End Sub

'Simple method that replaces completed outcome with Max Attempts
Sub Clean_Outcomes()
    If Current_Cell = "Completed" Then
        Current_Cell = "Max Attempts"
    End If
End Sub
'Removes any values found in debit day if the outcome is not "Confirmed"
Sub Clean_Debit_Day(currentRow As Integer)
    Dim Outcome_Cell As Object
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    If Outcome_Cell <> "Confirmed" And Outcome_Cell <> "Confirmed No Child" Then
        Current_Cell = Null
    End If
End Sub
'Puts the word "Monthly" in the frequency column for any record with an outcome of "Confirmed"
Sub Clean_Frequency(currentRow As Integer)
    Dim Outcome_Cell As Object
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child" Then
        Current_Cell = "Monthly"
    Else
        Current_Cell = Null
    End If
End Sub

'Looks for a shortened month name, then replaces that name with its long version.
'e.g. finds jan in "jan 2016" and replaces the cell with "January"
Sub Clean_Start_Month(rowNum As Integer)
    Dim Outcome_Cell As Object
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(rowNum, (Column_Number("Outcome")))
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child" Then
        'boolean value isMonth is set as True if a match is found for the short months
        isMonth = False
        'these are the values that the macro looks for in the start month cell.
        shortMonths = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
        'these are the values that the macro replaces in the start month
        longMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
        
        If Len(Trim(Current_Cell)) <> 0 Then
            Set_As_Text
            Proper_Case
            'loops over the short months to see if any of them are found in the start month cell.
            'if the match is made, the position of the value in the array shortMonths is stored in the variable index.
            For Each mon In shortMonths
                pos = InStr(Current_Cell, mon)
                If pos <> 0 Then
                    index = Application.Match(mon, shortMonths, False) - 1
                    'sets the value of the current cell to be longMonth value in the same position as the match in the ShortMonths array.
                    Current_Cell = longMonths(index)
                    'set to True if a month is found
                    isMonth = True
                End If
            Next mon
            'if no matches were made, then isMonth = false, and the cell needs to be highlighted yellow.
            If isMonth = False Then
                Highlight_Yellow
            End If
        'if the outcome is confirmed, and there is no start month, the cell needs to be highlighted yellow.
        Else
            Highlight_Yellow
        End If
    'if there is text in the start month and the outcome is not confirmed, the text is cleared.
    Else
        Current_Cell = Null
    End If
End Sub
'handles most variations of expiry dates being entered by agents.
Sub Clean_Expiry_Date()
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Text
        Current_Cell = Replace(Current_Cell, " ", "")
        pos = InStr(Current_Cell, "/")
        If pos <> 0 Then
            txtMonth = Split(Current_Cell, "/")(0)
            txtYear = Split(Current_Cell, "/")(1)
            If Len(Current_Cell) = 5 Then
                Current_Cell = txtMonth + "/" + "20" + txtYear
            ElseIf Len(Current_Cell) = 4 Then
                Current_Cell = "0" + txtMonth + "/" + "20" + txtYear
            ElseIf Len(Current_Cell) <> 7 Then
                Highlight_Yellow
            End If
        Else
            If Len(Current_Cell) = 4 Then
                Current_Cell = Mid(Current_Cell, 1, 2) + "/" + "20" + Mid(Current_Cell, 3, 2)
            ElseIf Len(Current_Cell) = 6 Then
                Current_Cell = Mid(Current_Cell, 1, 3) + "/" + Mid(Current_Cell, 2, 3)
            Else
                Highlight_Yellow
            End If
        End If
    End If
End Sub
'uses the luhn algorithm to check the validity of credit card numbers.
Sub Clean_Credit_Card_Number()
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Text
        Current_Cell = Replace(Current_Cell, " ", "")
        luhnSum = 0
        isDouble = False
        Length = Len(Current_Cell)
        For i = Length To 1 Step -1
            numStr = Mid(Current_Cell, i, 1)
            num = CInt(numStr)
            If (isDouble) Then
                newVal = num * 2
                If newVal > 9 Then
                    newVal = newVal - 9
                End If
                luhnSum = luhnSum + newVal
                isDouble = False
            Else
                luhnSum = luhnSum + num
                isDouble = True
            End If
        Next i
        If luhnSum Mod 10 <> 0 Then
            Highlight_Yellow
        End If
    End If
End Sub
'set the formatting to be currency if there is a value, and highlights the cell if there is no value.
Sub Clean_Single_Gift(rowNum As Integer)
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(rowNum, (Column_Number("Outcome")))
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Currency
    Else
        If Outcome_Cell = "Single Gift" Then
            Highlight_Yellow
        End If
    End If
End Sub

'set the formatting to be currency if there is a value, and highlights the cell if there is no value.
Sub Clean_Regular_Gift(rowNum As Integer)
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(rowNum, (Column_Number("Outcome")))
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Currency
        If Outcome_Cell <> "Confirmed" And Outcome_Cell <> "Confirmed No Child" Then
            Highlight_Yellow
        End If
    Else
        If Outcome_Cell = "Confirmed" Then
            Highlight_Yellow
        End If
    End If
End Sub
'formats the phone numbers, depending on if they are mobile numbers or home phone numbers.
'NOTE: This does not move the values to the correct columns
Sub Clean_Phone()
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Text
        Current_Cell = Replace(Current_Cell, " ", "")
        If Mid(Current_Cell, 1, 2) = "61" Then
            Current_Cell = Replace(Current_Cell, Mid(Current_Cell, 1, 2), "")
        End If
        
        If Left(Current_Cell, 1) <> "4" Then
            If Left(Current_Cell, 1) <> "0" Then
                Current_Cell = "0" & Current_Cell
            ElseIf Mid(Current_Cell, 2, 1) = "4" Then
                Current_Cell = Format(Current_Cell, "0000 000 000")
            End If
            Current_Cell = Format(Current_Cell, "00 0000 0000")
        Else
            If Mid(Current_Cell, 1, 1) <> "0" Then
                Current_Cell = Replace(Current_Cell, " ", "")
                Current_Cell = "0" & Current_Cell
                Current_Cell = Format(Current_Cell, "0000 000 000")
            Else
                Current_Cell = Replace(Current_Cell, " ", "")
                Current_Cell = Format(Current_Cell, "0000 000 000")
            End If
        End If
    End If
End Sub

Sub Clean_Postcode()
    If Len(Trim(Current_Cell)) <> 0 Then
        Set_As_Text
        firstDig = Mid(Current_Cell, 1, 1)
        If Mid(Current_Cell, 1, 1) = 8 And Len(Trim(Current_Cell)) = 3 Then
            Current_Cell = "0" & Current_Cell
        End If
        If Len(Trim(Current_Cell)) <> 4 Then
            Highlight_Yellow
        End If
    End If
End Sub

Sub Clean_State()
    If Len(Trim(Current_Cell)) = 0 Then
        Highlight_Yellow
    Else
        Upper_Case
    End If
    Is_State
End Sub

Sub Is_State()
    states = Array("", "NSW", "SA", "VIC", "WA", "QLD", "TAS", "NT", "ACT")
    isState = False
    For Each ausState In states
        If ausState = Current_Cell Then
           isState = True
        End If
    Next ausState
    If isState = False Then
        Highlight_Yellow
    End If
End Sub

Sub Clean_Address1()
    If Len(Trim(Current_Cell)) <> 0 Then
        Proper_Case
        If Len(Trim(Current_Cell)) < 10 Or Len(Trim(Current_Cell)) > 30 Then
            Highlight_Yellow
        End If
    Else
        Highlight_Yellow
    End If
End Sub

Sub Clean_Suburb()
    If Len(Trim(Current_Cell)) = 0 Then
        Highlight_Yellow
    Else
        Upper_Case
    End If
End Sub
'basic cleaning of gender. sets the value to be proper case, replaces M or F with male and female respectively
'or sets the value to be nothing
Sub Clean_Gender()
    If Len(Trim(Current_Cell)) <> 0 Then
        Proper_Case
        If Trim(Current_Cell) = "F" Or Trim(Current_Cell) = "Female" Then
            Current_Cell = "Female"
        ElseIf Trim(Current_Cell) = "M" Or Trim(Current_Cell) = "Male" Then
            Current_Cell = "Male"
        Else
            Current_Cell = Null
        End If
    End If
End Sub

Sub Clean_DOB()
    If Len(Trim(Current_Cell)) <> 0 Then
        Current_Cell.NumberFormat = "@"
        Current_Cell = Replace(Current_Cell, ".", "/")
        Current_Cell.NumberFormat = "@"
        Current_Cell = Format(Current_Cell, "DD/MM/YYYY")

    End If
End Sub

Sub Clean_Not_Null()
    If Len(Trim(Current_Cell)) = 0 Then
        Highlight_Yellow
    End If
End Sub

Sub Clean_Payment_Name()
    Proper_Case
End Sub

Sub Clean_Title()
    If Len(Trim(Current_Cell)) <> 0 Then
        Proper_Case
    End If
End Sub

Sub Clean_Name()
    If Len(Trim(Current_Cell)) = 0 Then
        Highlight_Yellow
    Else
        Proper_Case
    End If
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|---------------------------------------------END OF DATA CLEANING---------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-----------------------------------------GENERATE COLUMNS FROM EXISTING DATA----------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH GENERAL DATA GENERATION THAT IS GENERAL TO ALL CAMPAIGNS            |
'|    E.G PAYMENT DETAILS, GENERATION OF REPORTING COLUMNS                                                            |
'|____________________________________________________________________________________________________________________|

'Parent method for the procedures found below.
Sub Generate_Core_Data()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Clean_Payment_Details (currentRow)
        On Error GoTo HandleError
        Generate_Batch_month (currentRow)
        Populate_Call_Month (currentRow)
        Generate_Contacts_And_Invalid (currentRow)
        Generate_Call_Week (currentRow)
        Generate_No_Phone (currentRow)
    Next currentRow
Exit Sub
HandleError:
    Handle_Core_Data_Failure
    Exit Sub
End Sub

Sub Generate_STC_Core_Data()
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Clean_Payment_Details (currentRow)
        On Error GoTo HandleError
        'Generate_Batch_month (currentRow)
        'Populate_Call_Month (currentRow)
        'Generate_Contacts_And_Invalid (currentRow)
        'Generate_Call_Week (currentRow)
        'Generate_No_Phone (currentRow)
    Next currentRow
Exit Sub
HandleError:
    Handle_Core_Data_Failure
    Exit Sub
End Sub

'Parent method to clean payment details for ALL charities. Calls the below methods.
Sub Clean_Payment_Details(currentRow As Integer)
        Initialize_Payment_Row (currentRow)
        Validate_Payment_Method
        Validate_Card_Type
        Validate_Card_Name
        Validate_Expiry
End Sub

'Initialized all the cells needed to clean the payment details. Note: This procedure contains
'an error handler which will determine where exactly any initialization issues occur.
'Errors in this section will almost always be due to mismatched column names
Sub Initialize_Payment_Row(currentRow As Integer)
    On Error GoTo HandleError
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Payment_Method_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
    Set Card_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
    Set Card_Number_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
    Set Card_Expiry_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
    Set Card_Name_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Name on Card")))
    Set BSB_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BSB")))
    Set Bank_Name_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Bank Name")))
    Set Account_Num_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Account Number")))
    Set Account_Name_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Account Holders Name")))
Exit Sub
HandleError:
    Handle_Initialize_Failure
    Exit Sub
End Sub

Sub Validate_Expiry()
    If Payment_Method_Cell = "Credit Card" Then
        If Len(Trim(Card_Expiry_Cell)) = 0 Then
            Card_Expiry_Cell.Interior.ColorIndex = 6
        End If
    End If
End Sub

'Ensures the correct payment method was selected by the agent for the payments details provided.
Sub Validate_Payment_Method()
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Single Gift" Or Outcome_Cell = "Confirmed No Child" Then
        If Len(Trim(Card_Number_Cell)) <> 0 Then
           Payment_Method_Cell = "Credit Card"
        ElseIf Len(Trim(BSB_Cell)) <> 0 Then
           Payment_Method_Cell = "Direct Debit"
        End If

    Else
        If Len(Trim(Payment_Method_Cell)) <> 0 Then
            Payment_Method_Cell = Null
        End If
    End If
End Sub

'Ensures card type matches the card number. eg. 4 -> visa, 5 -> mastercard etc.
Sub Validate_Card_Type()
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Single Gift" Or Outcome_Cell = "Confirmed No Child" Then
        If Payment_Method_Cell = "Credit Card" Then
            If Left(Card_Number_Cell, 1) = "4" Then
                Card_Type_Cell = "Visa"
            ElseIf Left(Card_Number_Cell, 1) = "5" Then
                Card_Type_Cell = "Mastercard"
            ElseIf Left(Card_Number_Cell, 1) = "3" Then
                Card_Type_Cell = "Amex"
            Else
               Card_Type_Cell.Interior.ColorIndex = 6
            End If
        ElseIf Payment_Method_Cell = "Direct Debit" Then
            Card_Type_Cell = Null
        End If
    Else
        Card_Type_Cell = Null
    End If
End Sub

'Puts the Card Name in proper case or highlights the cell if it is empty and shouldnt be.
Sub Validate_Card_Name()
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Single Gift" Or Outcome_Cell = "Confirmed No Child" Then
        If Len(Trim(Card_Name_Cell)) <> 0 Then
            Card_Name_Cell = StrConv(Card_Name_Cell, vbProperCase)
        Else
            If Len(Trim(Card_Number_Cell)) <> 0 Then
                Card_Name_Cell.Interior.ColorIndex = 6
            End If
        End If
    Else
        Card_Name_Cell = Null
    End If
End Sub

'Uses a regex pattern initialized in the beginning of the macro as 'listMatch" which looks
'for 6 straight numbers signifying a date. THIS DOES NOT WORK CORRECTLY FOR LISTS WITH NO DATE.
Sub Generate_Batch_month(currentRow As Integer)
    Dim List_Cell As Object
    Dim Batch_Cell As Object
    
    Column_Not_Found = "Batch Month"
    capitalMONTHS = Array("JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER")
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Set List_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DataSetName")))
    Set Batch_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Batch Month")))
    Set weekVal = ListMatch.Execute(List_Cell)
    isMonthName = False
    
    For Each capMonth In capitalMONTHS
        pos = InStr(List_Cell, capMonth)
        If pos <> 0 Then
            Batch_Cell = StrConv(capMonth, vbProperCase)
            isMonthName = True
        End If
    Next capMonth
    If isMonthName <> True Then
    'On Error Resume Next
        If weekVal.Count <> 0 Then
            monthVal = Mid(weekVal(0), 3, 2)
            Batch_Cell = longMonths(CInt(monthVal))
        End If
    End If
End Sub

'Sets the call month by pulling the month from OutcomeUpdateDateTime
Sub Populate_Call_Month(currentRow As Integer)
    Dim OutcomeUpdate_Cell As Object
    Dim Call_Month_Cell As Object
    
    Column_Not_Found = "Call Month"
    
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    Set Call_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Call Month")))
    
    Call_Month_Cell = Format(OutcomeUpdate_Cell, "MMMM")
End Sub

'Sets contacts and invalids in the weekly reporting section of the data.
Sub Generate_Contacts_And_Invalid(currentRow As Integer)
    Dim Contact_Cell As Object
    Dim Invalid_Cell As Object
    Dim Outcome_Cell As Object
    
    Column_Not_Found = "Contacts vs Invalid"

    Set Contact_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Contact")))
    Set Invalid_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Invalid")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        
    If (Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child" Or Outcome_Cell = "Single Gift" Or Outcome_Cell = "Not Interested" _
    Or Outcome_Cell = "Instant Refusal" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Promised RG" _
    Or Outcome_Cell = "Promised SG" Or Outcome_Cell = "Already a Supporter" Or Outcome_Cell = "Do Not Call" _
    Or Outcome_Cell = "Deceased" Or Outcome_Cell = "Already a supporter" Or Outcome_Cell = "Already Updated Details" _
    Or Outcome_Cell = "Previously Cancelled" Or Outcome_Cell = "Update Details Only" Or Outcome_Cell = "Cancelled" _
    Or Outcome_Cell = "Updated Details Only") Then
        Contact_Cell = "Contact"
    ElseIf Outcome_Cell = "Completed" Or Outcome_Cell = "Max Attempts" Then
        Outcome_Cell = "Max Attempts"
    Else
        Invalid_Cell = "Yes"
    End If
End Sub

'Generates the Week Of Data Column
Sub Generate_Call_Week(currentRow As Integer)
    Dim Call_Week_Cell As Object
    
    Column_Not_Found = "Week Of Data"

    Temp1 = Workbooks("WAYS Data Cleaning Macro").Sheets("Macro").Cells(4, "B")
    Temp1 = Format(Temp1, "dd/mm/yy")
    Temp2 = DateAdd("d", -6, Temp1)
    Temp2 = Format(Temp2, "dd/mm/yy")
    
    Set Call_Week_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Week Of Data")))
    
    Call_Week_Cell = Temp2 & " - " & Temp1
End Sub

'Contains logic to switch the mobile and home phone numbers if they are in the wrong column
Sub Fix_Home_Phone_And_Mobile_Phone()
    Dim Home_Cell As Object
    Dim Mobile_Cell As Object

    Column_Not_Found = "Fix Home and Mobile Phone"
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Home_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("HomePhone")))
        Set Mobile_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("MobilePhone")))
    
        If Mid(Mobile_Cell, 2, 1) <> "4" And Len(Trim(Home_Cell)) = 0 Then
            Home_Cell = Mobile_Cell
            Mobile_Cell = Null
        End If
        
        If Mid(Home_Cell, 2, 1) = "4" And Len(Trim(Mobile_Cell)) = 0 Then
            Mobile_Cell = Home_Cell
            Home_Cell = Null
        End If
    
        If Mid(Home_Cell, 2, 1) = "4" And Mid(Mobile_Cell, 2, 1) <> "4" Then
            temp = Home_Cell
            Home_Cell = Mobile_Cell
            Mobile_Cell = temp
        End If
    
        If Len(Home_Cell) <> 0 And Mid(Home_Cell, 2, 1) <> "4" And Mid(Mobile_Cell, 2, 1) <> "4" Then
            Mobile_Cell = Null
        End If
    
        If Home_Cell = Mobile_Cell Then
            Home_Cell = Null
        End If
    Next currentRow
End Sub

'Sets "No" on no phone contact for calls dispositioned as Do Not Call
Sub Generate_No_Phone(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim NoPhone_Cell As Object
    
    Column_Not_Found = "No Phone"
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set NoPhone_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
    
    If Outcome_Cell = "Do Not Call" Then
        NoPhone_Cell = "No"
    End If

End Sub

'Sets the Cluster to whatever string is passed in as a parameter for each row.
Sub Generate_Cluster(Cluster As String)
Dim Cluster_Cell As Object
clusterCol = Column_Number("Cluster")
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Cluster_Cell = Worksheets(Initiative_Name).Cells(currentRow, clusterCol)
    Cluster_Cell = Cluster
Next currentRow
End Sub
Sub Set_Acq_Source_To(Source As String)
Dim Acq_Source As Object
sourceCol = Column_Number("Acq Source")
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Acq_Source = Worksheets(Initiative_Name).Cells(currentRow, sourceCol)
    Acq_Source = Source
Next currentRow

End Sub

'Sets the Type to whatever string is passed in as a parameter for each row.
Sub Generate_Type(initType As String)
    Dim Type_Cell As Object
    typeCol = Column_Number("Type")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, typeCol)
        Type_Cell = initType
    Next currentRow
End Sub

'Loops over shortSources, and determines whether a shortSource exists in the list name.
'If it does, it uses longSources to generate the correct name.
Sub Generate_Acq_Source_Lead_Conversion()
Dim List_Cell As Object
Dim Acq_Cell As Object

'what we look for in the list name
shortSources = Array("cohort grid", "cohort s", "3di", "omni", "offers", "cause", _
"marketing", "egentic", "opentop", "8th", "quinn", "zinq", "upside", "rokt", _
"contact", "vizmond", "kobi", "empowered", "change.org", "dataphoria", "luna", "c7")

'what we use to put in the acq source column
LongSources = Array("Cohort - grid", "Cohort - Stand Alone", "3Di - Stand Alone", _
"Omni - Phone Leads", "Offers Now - Stand Alone", "Cohort - Cause Grid", _
"Marketing Punch - Stand Alone", "eGentic - Stand Alone", "OpenTop - Stand Alone", _
"8th Floor - Stand Alone", "Quinn - Phone Leads", "Zinq - Stand Alone", _
"Upside - Stand Alone", "ROKT - Stand Alone", "Contact Me - Stand Alone", _
"Vizmond - Stand Alone", "Kobi - Stand Alone", "Empowered - Stand Alone", _
"Change.org", "Dataphoria", "Luna Park - Stand Alone", "C7 - Stand Alone")

For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set List_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DataSetName")))
    Set Acq_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Acq Source")))
    isSource = False
    temp = StrConv(List_Cell, vbLowerCase)
    For Each Source In shortSources
        pos = InStr(temp, Source)
        If pos <> 0 Then
            index = Application.Match(Source, shortSources, False) - 1
            Acq_Cell = LongSources(index)
            isSource = True
        End If
    Next Source
    If isSource = False Then
        Acq_Cell.Interior.ColorIndex = 6
    End If
Next currentRow
End Sub

'For all rows with a title, but no gender -> generates gender.
'For all rows with a gender, but no title -> generates title.
Sub Populate_Title_And_Gender()
    genderCol = Column_Number("Gender")
    titleCol = Column_Number("Title")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Gender_Cell = Worksheets(Initiative_Name).Cells(currentRow, genderCol)
        Set Title_Cell = Worksheets(Initiative_Name).Cells(currentRow, titleCol)
        If Len(Trim(Title_Cell)) = 0 Then
           Generate_Title_From_Gender
        ElseIf Len(Trim(Gender_Cell)) = 0 Then
                Generate_Gender_From_Title
        End If
    Next currentRow
End Sub

'If Gender cell is not empty, sets title cell based on it's value.
Sub Generate_Title_From_Gender()
    If Gender_Cell = "Female" Then
        Title_Cell = "Ms"
    ElseIf Gender_Cell = "Male" Then
        Title_Cell = "Mr"
    End If
End Sub

'If title cell is not empty, sets gender cell based on it's value.
Sub Generate_Gender_From_Title()
    If Title_Cell = "Mr" Then
        Gender_Cell = "Male"
    ElseIf Title_Cell = "Ms" Or Title_Cell = "Mrs" Or Title_Cell = "Miss" Then
        Gender_Cell = "Female"
    End If
End Sub

Sub Generate_CallDate(calldate As String)
    Dim OutcomeUpdate_Cell As Object
    Dim Call_Date_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Call_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number(calldate)))
        Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
  
        Call_Date_Cell.NumberFormat = "@"
        Call_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    Next currentRow
End Sub

Sub Generate_Upgrade_Amount()
    Dim Outcome As Object
    Dim Existing_RG As Object
    Dim New_RG As Object
    Dim Upgrade As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Existing_RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Current Gift Amount")))
        Set New_RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
        Set Upgrade = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Upgrade Amount")))
        If Outcome = "Confirmed" Then
            If Len(Trim(New_RG)) > 0 Then
                Upgrade = New_RG - Existing_RG
                If Upgrade <= 0 Then
                    Upgrade.Interior.ColorIndex = 6
                End If
            Else
                Upgrade.Interior.ColorIndex = 6
            End If
        End If
        
    Next currentRow
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF COLUMN GENERATION-------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-----------------------------------------CHARITY/CAMPAIGN SELECTION-------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH SELECTING THE CORRECT CLEANING/GENERATION METHODS FOR EACH          |
'|    SPECIFIC CHARITY AND INITITIAVE. NO ACTUAL CLEANING / GENERATION OCCURS IN THIS SECTION, BUT RATHER THE         |
'|    LOGIC REGARDING HOW EACH CHARITY AND INITIATIVE NEEDS TO BE CLEANED                                             |
'|____________________________________________________________________________________________________________________|

'Selects data generation depending on the name of the charity selected
Sub Choose_Data_Generation()
    Select Case Charity
        Case "WAP"
            Select_Data_Generation_WAP
        Case "AIA"
            Select_Data_Generation_AIA
        Case "Taronga"
            Select_Data_Generation_Taronga
        Case "CMRI"
            Select_Data_Generation_CMRI
        Case "Northcott"
            Select_Data_Generation_Northcott
        Case "MAW"
            Select_Data_Generation_MAW
        Case "CCIA"
            Select_Data_Generation_CCIA
        Case "ICV"
            Select_Data_Generation_ICV
        Case "TSC"
            Select_Data_Generation_TSC
        Case "Salvos"
            Select_Data_Generation_Salvos
        Case "HF"
            Select_Data_Generation_HF
        Case "BH"
            Select_Data_Generation_BH
        Case "TSF"
            Select_Data_Generation_TSF
        Case "WWF"
            Select_Data_Generation_WWF
        Case "PM"
            Select_Date_Generation_Pmac
        Case "Starlight"
            Select_Data_Generation_Starlight
        Case "Opportunity"
            Select_Data_Generation_Opportunity
        Case "H4H"
            Select_Data_Generation_H4H
        Case "Wesley"
            Select_Data_Generation_Wesley
        Case "RedKite"
            Select_Data_Generation_RedKite
        Case "Greenpeace"
            Select_Data_Generation_Greenpeace
        Case "Variety"
            Select_Data_Generation_Variety
        Case "STC"
            Select_Data_Generation_STC
        Case "WaterAid"
            Select_Data_Generation_WaterAid
        Case "Mission"
            Select_Data_Generation_Mission
        Case "UN-Women"
            Select_Data_Generation_UNWomen
        Case "MSF"
            Select_Data_Generation_MSF
        Case "AMF"
            Select_Data_Generation_AMF
        Case "CPA"
            Select_Data_Generation_CPA
    End Select
End Sub
Sub Select_Data_Generation_CPA()
 Select Case Initiative_Type
        Case "Recycled"
            Generate_Acq_Source_Lead_Conversion
            Generate_Cluster ("Lead Conversion")
            Generate_Cerebral_Palsy
    End Select
End Sub
Sub Select_Data_Generation_AMF()
    Select Case Initiative_Type
        Case "Lead Conversion"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
        Case "Reactivations"
            Generate_Type ("Reactivations")
            Generate_Cluster ("Reactivations")
        Case "Upgrades"
            Generate_Type ("Upgrades")
            Generate_Cluster ("Upgrades")
            Generate_Upgrade_Amount
            Generate_Report_Upgrade_Amount_ICV
    End Select

End Sub

Sub Select_Data_Generation_MSF()
Select Case Initiative_Type
    Case "Mid Donor Thankyou"
        Generate_Type ("Thankyou")
        Generate_Cluster ("Thankyou")
End Select
End Sub

Sub Select_Data_Generation_UNWomen()
Select Case Initiative_Type
    Case "Cash Conversion"
        Generate_Type ("Cash Conversion")
        Generate_Cluster ("Cash Conversion")
End Select
End Sub

Sub Select_Data_Generation_Mission()
Select Case Initiative_Type
    Case "Reactivations"
       Generate_Mission_Reactivations
    Case "Upgrades"
       Generate_Mission_Upgrades
End Select
End Sub

Sub Select_Data_Generation_WaterAid()
Select Case Initiative_Type
    Case "Upgrades"
        Populate_Title_And_Gender
        Generate_Type ("Upgrades")
        Generate_Cluster ("Upgrades")
        Generate_Upgrade_Amount_WaterAid
        Generate_Call_Date_WaterAid
End Select
End Sub
Sub Select_Data_Generation_STC()
Select Case Initiative_Type
    Case "Inbound"
        Generate_STC
    Case "Reactivation"
        Generate_Debit_Date
        Generate_Cluster ("Reactivation")
        Generate_STC_Reacts
        
        
End Select
End Sub

Sub Select_Data_Generation_Variety()
Select Case Initiative_Type
    Case "Lead Conversion"
        Generate_Data_Lead_Conversion
End Select
End Sub

Sub Select_Data_Generation_Greenpeace()
Select Case Initiative_Type
    Case "Lead Conversion"
        Generate_Acq_Source_Lead_Conversion
        Generate_Cluster ("Lead Conversion")
        Populate_Title_And_Gender
        Greenpeace_Generation
    Case "Quinn"
        Generate_Acq_Source_Lead_Conversion
        Generate_Cluster ("Lead Conversion")
        Populate_Title_And_Gender
        Greenpeace_Generation
    Case "DRTV"
        Generate_Acq_Source_Lead_Conversion
        Generate_Cluster ("DRTV")
        Populate_Title_And_Gender
        Greenpeace_DRTV_Gen
    Case "ReachTel"
        Generate_Acq_Source_Lead_Conversion
        Generate_Cluster ("Lead Conversion")
        Populate_Title_And_Gender
        Greenpeace_Generation
    Case "OpenTop"
        Generate_Acq_Source_Lead_Conversion
        Generate_Cluster ("Lead Conversion")
        Populate_Title_And_Gender
        Greenpeace_Generation
End Select

End Sub

Sub Select_Data_Generation_RedKite()
Select Case Initiative_Type
    Case "Upgrades"
        Generate_Data_Lead_Conversion
    Case "Supporter Conversion"
        Generate_Data_Lead_Conversion
End Select

End Sub

Sub Select_Data_Generation_Wesley()
Select Case Initiative_Type
    Case "Reactivations"
        Generate_Data_Lead_Conversion
        Populate_Title_And_Gender
        Format_Expiry_Date
    Case "Declines"
        Generate_Data_Lead_Conversion
        Populate_Title_And_Gender
        Format_Expiry_Date
End Select


End Sub

Sub Select_Data_Generation_H4H()
Select Case Initiative_Type
    Case "Lead Conversion"
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
        Populate_Title_And_Gender
End Select



End Sub
Sub Select_Data_Generation_Opportunity()
Select Case Initiative_Type
    Case "Lead Conversion"
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
        Populate_Title_And_Gender
        Opportunity_Generation
End Select

End Sub

Sub Select_Data_Generation_Starlight()
Select Case Initiative_Type
    Case "Lead Conversion"
        Starlight_Generation
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
        Populate_Title_And_Gender
End Select


End Sub

Sub Select_Date_Generation_Pmac()
Select Case Initiative_Type
    Case "Lead Conversion"
        Populate_Title_And_Gender
        Generate_Peter_Mac
End Select
End Sub


Sub Select_Data_Generation_WWF()
Select Case Initiative_Type
        Case "Lead Conversion"
            WWF_Generation
            Generate_Data_Lead_Conversion
         Case "Recycled"
            WWF_Generation
            Generate_Cluster ("Recycled")
         Case "Cash Conversion"
            WWF_Generation
            Generate_Data_Lead_Conversion
         Case "Tiger Petition"
            WWF_Petition_Generation
            Generate_Cluster ("Petition")
            Set_Acq_Source_To ("Tiger")
         Case "Koala Petition"
            WWF_Petition_Generation
            Generate_Cluster ("Petition")
            Set_Acq_Source_To ("Koala")
         Case "FFTR Petition"
            WWF_Petition_Generation
            Set_Acq_Source
            Generate_Cluster ("Petition")
    End Select

End Sub

Sub Select_Data_Generation_TSF()
Select Case Initiative_Type
        Case "Lead Conversion"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Generate_Country
            Generate_Agency
            Generate_Signup_Date
            Generate_Debit_Date
            Generate_Payment_method
            Convert_Expiry_to_two_digits
            Generate_DonationType_SmithFamily
          
        Case "ReachTel"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Generate_Country
            Generate_Agency
            Generate_Signup_Date
            Generate_Debit_Date
            Generate_Payment_method
            Convert_Expiry_to_two_digits
            Generate_DonationType_SmithFamily
          
        Case "Bank Rejects"
            Generate_Cluster ("Bank Rejects")
            Generate_ArrearsAmount_TSF_Rejects
            Copy_ConsumerId_To_Supporter_ID_TSF_Rejects
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Generate_Country
            Generate_Agency
            Generate_Signup_Date
            Generate_Debit_Date
            Generate_Payment_method
            Convert_Expiry_to_two_digits
            Generate_DonationType_SmithFamily
    End Select

End Sub
Sub Select_Data_Generation_BH()
Select Case Initiative_Type
        Case "Recycled"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
    End Select

End Sub

Sub Select_Data_Generation_HF()
    Select Case Initiative_Type
        Case "Lead Conversion"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
         Case "Bequest"
            Generate_HF_Bequests
             Generate_Cluster ("Bequest")
    End Select
End Sub

Sub Select_Data_Generation_Salvos()
    Select Case Initiative_Type
        Case "Lead Conversion"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Check_Token_Card_Salvos
        Case "Recycled"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Check_Token_Card_Salvos
        Case "SG Experian"
            Generate_Data_Lead_Conversion
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            Check_Token_Card_Salvos
    End Select
End Sub

Sub Select_Data_Generation_TSC()
 Select Case Initiative_Type
    Case "Lead Conversion"
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone

 End Select
End Sub

Sub Select_Data_Generation_ICV()
    Select Case Initiative_Type
        Case "Lead Conversion"
            Generate_Acq_Source_Lead_Conversion
            Generate_Type ("Acquisition")
            Generate_Cluster ("Lead Conversion")
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            
        Case "Recycled"
            Generate_Acq_Source_Lead_Conversion
            Generate_Type ("Acquisition")
            Generate_Cluster ("Recycled")
            Fix_Home_Phone_And_Mobile_Phone
            Populate_Title_And_Gender
            
        Case "C2C Insights"
            Populate_Title_And_Gender
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Type ("Cash Conversion")
            Generate_Cluster ("C2C")
        Case "Cash Conversion"
            Populate_Title_And_Gender
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Type ("Cash Conversion")
            
        Case "Upgrades"
            Populate_Title_And_Gender
            Generate_Type ("Upgrades")
            Generate_Cluster ("Upgrades")
            WAP_Generate_Upgrade_Amount
        Case "Reactivations"
            Generate_Cluster ("Reactivations")
    End Select
End Sub

Sub Select_Data_Generation_CCIA()
    Select Case Initiative_Type
        'Handles MAW Recycled
        Case "Lead Conversion"
             Generate_Data_Lead_Conversion
    End Select


End Sub

'Handles all data generation logic for Make a Wish
Sub Select_Data_Generation_MAW()
    Select Case Initiative_Type
        'Handles MAW Recycled
        Case "Lead Conversion"
            Generate_Data_Lead_Conversion
            Generate_MAW
         Case "Recycled"
            Generate_Cluster ("Recycled")
            Generate_MAW
            
    End Select
End Sub

'Handles all data generation logic for Northcott
Sub Select_Data_Generation_Northcott()
Select Case Initiative_Type

    'Handles Northcott Lead Conversion
    Case "Lead Conversion"
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
End Select
End Sub

'Handles all data generation logic for Taronga
Sub Select_Data_Generation_Taronga()
Select Case Initiative_Type

    'Handles Taronga Lead Conversion
    Case "Lead Conversion"
        Populate_Title_And_Gender
        Generate_Acq_Source_Lead_Conversion
        Generate_Type ("Acquisition")
        Generate_Cluster ("Lead Conversion")
        Fix_Home_Phone_And_Mobile_Phone
    
    Case "LC Recycled"
        Populate_Title_And_Gender
        Generate_Acq_Source_Lead_Conversion
        Generate_Type ("Acquisition")
        Generate_Cluster ("Recycled")
        Fix_Home_Phone_And_Mobile_Phone
    
    Case "Cash Conversion"
        Populate_Title_And_Gender
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Type ("Cash Conversion")
        Generate_Cluster ("C2C")
        
    Case "Declines"
        Generate_CallDate ("Last Call Date")
        Generate_Cluster ("Declines")
        Format_Expiry_Date
        
    Case "Reactivations"
        Populate_Title_And_Gender
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Type ("Reactivations")
        Generate_Cluster ("Reactivations")
    
    Case "Bilby Adoption"
        Populate_Title_And_Gender
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Cluster ("Bibly Adoption")
        
    Case "Upgrades"
        Populate_Title_And_Gender
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Upgrade_Amount
        Generate_Type ("Upgrades")
        Generate_Cluster ("Upgrades")
End Select
End Sub

'Handles all data generation logic for CMRI
Sub Select_Data_Generation_CMRI()
Select Case Initiative_Type
    
    'Handles CMRI Lead Conversion
    Case "Online"
        Generate_Data_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
    
    'Handles CMRI Recycled
    Case "Recycled"
        Generate_Data_CMRI_Recycled
        Fix_Home_Phone_And_Mobile_Phone
End Select
End Sub

'Handles all data generation logic for Amnesty
Sub Select_Data_Generation_AIA()
Select Case Initiative_Type

    'Handles AIA Lead Conversion
    Case "Online Survey"
        Generate_Acq_Source_Lead_Conversion
        Generate_Type ("Online")
        Generate_Cluster ("Acquisition")
        Generate_AIA_Lead_Conversion
        Fix_Home_Phone_And_Mobile_Phone
        Populate_Title_And_Gender
        
    'Handles AIA Petition Conversion
    Case "Petition Conversion"
        Generate_AIA_Petition_Conversion
        Generate_Data_AIA_Petiton_Converion
        Fix_Home_Phone_And_Mobile_Phone
        
    'Handles AIA Online Recycle
    Case "Online Recycle"
        Generate_AIA_Lead_Conversion
        Populate_Title_And_Gender
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Cluster ("Online Recycle")
        Generate_Acq_Source_Recycled
        Set_Campaign ("Online Recycle")
        
    'Handles AIA Cash Conversion
    Case "Cash Conversion"
        Generate_AIA_Cash_Conversion
        Fix_Home_Phone_And_Mobile_Phone
        Generate_Cluster ("Cash Conversion")
End Select
End Sub


'Handles all data generation logic for World Animal Protection
Sub Select_Data_Generation_WAP()
    Select Case Initiative_Type
    
        'Handles WAP Lead Conversion
        Case "Lead Conversion"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Acq_Source_Lead_Conversion
            Generate_Cluster ("Lead Conversion")
            Populate_Title_And_Gender
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
        Case "Reach Tel"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Acq_Source_Lead_Conversion
            Generate_Cluster ("Lead Conversion")
            Populate_Title_And_Gender
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            
        'Handles WAP Internal Petition
        Case "Internal Petition"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Data_Internal_Petition
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
        
         Case "Dolphin Petition"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Data_Internal_Petition
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
        
         Case "Petition Recycled"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Data_Internal_Petition
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            
         Case "Paid Lead Recycled"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Acq_Source_Lead_Conversion
            Generate_Cluster ("Lead Conversion")
            Populate_Title_And_Gender
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
        'Handles WAP Upgrade Frtnightly Rolling
        Case "Upgrades"
            WAP_Generate_Start_Month_Warm
            WAP_Generate_Import_Result
            Generate_Cluster ("Upgrades")
            WAP_Generate_Upgrade_Amount
            
        'Handles WAP Upgrade Monthly Rolling
        Case "Upgrade Monthly Rolling"
            WAP_Generate_Start_Month_Warm
            WAP_Generate_Import_Result
            Generate_Type ("Rolling")
            Generate_Cluster ("Upgrades")
            WAP_Generate_Upgrade_Amount
            
        'Handles WAP SC Jan13 - Dec14
        Case "SC Jan-Dec"
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Type ("Catch Up")
            Generate_Cluster ("Supporter Conversion")
            
        Case "Supporter Conversion"
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Cluster ("Supporter Conversion")
            
        'Handles WAP Bank Rejects
        Case "Bank Rejects"
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Cluster ("Bank Rejects")
            Generate_Outcomes_Bank_Rejects
            
        'Handles WAP Lapsed Catch Up"
        Case "Lapsed Catch Up"
            WAP_Lapsed_Generate_Debit_Date
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Type ("Catch Up")
            Generate_Cluster ("Lapsed")
            
        'Handles WAP Lapsed Rolling"
        Case "Lapsed Rolling"
            WAP_Lapsed_Generate_Debit_Date
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Type ("Rolling")
            Generate_Cluster ("Lapsed")
            
        Case "Long Lapsed"
            WAP_Lapsed_Generate_Debit_Date
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Cluster ("Long Lapsed Reactivation")
            
        Case "Lapsed 2015 & New"
            WAP_Lapsed_Generate_Debit_Date
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Cluster ("Lapsed Reactivation")
            
        Case "Wave 2 Lead Con Recycle"
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
            Generate_Type ("Online")
            Generate_Cluster ("Recycle")
            
        Case "Thankyou"
            WAP_Generate_Start_Month_Acqusition
            Generate_Cluster ("Thank You")
            
        Case "Quinn"
            Fix_Home_Phone_And_Mobile_Phone
            Generate_Acq_Source_Lead_Conversion
            Generate_Cluster ("Lead Conversion")
            Populate_Title_And_Gender
            WAP_Generate_Start_Month_Acqusition
            WAP_Generate_Import_Result
    End Select
End Sub

Sub Generate_HF_Bequests()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        HF_Move_New_Phones (currentRow)
    Next currentRow
    
End Sub

Sub HF_Move_New_Phones(currentRow As Integer)
    Dim newHome As Object
    Dim newMobile As Object
    Dim newPhone1 As Object
    Dim newPhone2 As Object
    
    Set newHome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New Home")))
    Set newMobile = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New Mobile")))
    Set newPhone1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New Phone Num1")))
    Set newPhone2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New Phone Num2")))
    
    If Len(Trim(newMobile)) = 0 Then
        
        If Len(Trim(newPhone1)) > 0 Then
        
            If Mid(newPhone1, 2, 1) = "4" Then
                newMobile = newPhone1
            End If
      End If
        If Len(Trim(newPhone2)) > 0 Then
        
            If Mid(newPhone2, 2, 1) = "4" And Len(Trim(newMobile)) = 0 Then
                newMobile = newPhone2
            End If
        End If
    End If
    
    If Len(Trim(newHome)) = 0 Then
        
        If Len(Trim(newPhone1)) > 0 Then
        
            If Mid(newPhone1, 2, 1) <> "4" Then
                newHome = newPhone1
            End If
        End If
        If Len(Trim(newPhone2)) > 0 Then
        
            If Mid(newPhone2, 2, 1) <> "4" And Len(Trim(newHome)) = 0 Then
                newHome = newPhone2
            End If
        End If
    End If
    
End Sub

'Generates AIA Petition Specific Data
Sub Generate_Data_AIA_Petiton_Converion()
    Generate_Cluster ("Petition Conversion")
    Generate_AIA_Petition_Acq_Source ("Petition Conversion")
End Sub

'Generates CMRI Recycled specific data
Sub Generate_Data_CMRI_Recycled()
    Populate_Title_And_Gender
    Generate_Cluster ("Recycled")
End Sub

'Generates General lead conversion data
Sub Generate_Data_Lead_Conversion()
   
    Generate_Acq_Source_Lead_Conversion
    Generate_Type ("Online")
    Generate_Cluster ("Acquisition")
    Populate_Title_And_Gender
End Sub

'Generates WAP Internal Petition specific data
Sub Generate_Data_Internal_Petition()
    Populate_Title_And_Gender
    Generate_Cluster ("Petition Conversion")
End Sub

Sub Generate_Cerebral_Palsy()
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Generate_Static_RG_Fields_Cerebral_Palsy (currentRow)
    Generate_Static_SG_Fields_Cerebral_Palsy (currentRow)
    Set_Card_Type_Cerebral_Palsy (currentRow)
    Change_Contact_OptOuts_To_Y_Cerebral_Palsy (currentRow)
    Add_SG_To_RG_Col_And_Fix_RG (currentRow)
Next currentRow
Generate_CallDate ("Source Date")
Generate_Debit_Date
End Sub

Sub Add_SG_To_RG_Col_And_Fix_RG(currentRow As Integer)
    Dim Installment As Object
    Dim SG As Object
    Dim RG As Object
    Dim Outcome As Object
    
    Set Installment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Installment")))
    Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG Amount")))
    Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    If Len(Trim(Installment)) > 0 Then
        RG = Installment
        RG = Format(RG, "Currency")
        If Outcome <> "Confirmed" Then
            RG.Interior.ColorIndex = 6
        Else
            RG.Interior.ColorIndex = 0
        End If
    
    End If
    
    If Len(Trim(SG)) > 0 Then
        Installment = SG
        Installment = Format(Installment, "Currency")
        If Outcome <> "Single Gift" Then
            SG.Interior.ColorIndex = 6
        Else
            SG.Interior.ColorIndex = 0
        End If
    End If
End Sub

Sub Change_Contact_OptOuts_To_Y_Cerebral_Palsy(currentRow As Integer)
    Dim NoMail As Object
    Dim NoEmail As Object
    Dim NoPhone As Object
    
    Set NoMail = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Mail")))
    Set NoEmail = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Email")))
    Set NoPhone = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
    
    If Len(Trim(NoMail)) > 0 Then
        NoMail = "Y"
    End If
    
    If Len(Trim(NoEmail)) > 0 Then
        NoEmail = "Y"
    End If
    
    If Len(Trim(NoPhone)) > 0 Then
        NoPhone = "Y"
    End If

End Sub
Sub Set_Card_Type_Cerebral_Palsy(currentRow As Integer)
    Dim CardType As Object
    Dim CardNumber As Object
    Dim Outcome As Object
    Dim PaymentMethod As Object
    
    Set CardType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
    Set PaymentMethod = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
    Set CardNumber = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    If (Outcome = "Confirmed" Or Outcome = "Single Gift") And PaymentMethod = "Credit Card" Then
        If Mid(CardNumber, 1, 1) = "4" Then
            CardType = "VISA"
        ElseIf Mid(CardNumber, 1, 1) = "5" Then
            CardType = "MCARD"
        ElseIf Mid(CardNumber, 1, 1) = "3" Then
            CardType = "AMEX"
        Else
            CardType.Interior.ColorIndex = 6
        End If
    End If
End Sub

Sub Generate_Static_RG_Fields_Cerebral_Palsy(currentRow As Integer)
    Dim ContactType As Object
    Dim PledgePlan As Object
    Dim PledgeType As Object
    Dim Method As Object
    Dim ReceiptSummary As Object
    Dim ReceiptReq As Object
    Dim OneOff As Object
    Dim CRMPrimaryManager As Object
    Dim Param1Name As Object
    Dim SourceCode As Object
    Dim Outcome As Object
    Dim State As Object
    Dim StatementText As Object
    Dim Param2Value As Object
    Dim Source As Object
    
    Set ContactType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Contact Type")))
    Set PledgePlan = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Pledge Plan")))
    Set PledgeType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Pledge Type")))
    Set Method = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Method")))
    Set ReceiptSummary = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Receipt Summary")))
    Set ReceiptReq = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Receipt Required?")))
    Set OneOff = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("One Off?")))
    Set CRMPrimaryManager = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CRM Primary Manager")))
    Set Param1Name = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Parameter 1 Name")))
    Set SourceCode = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Source Code")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set State = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("State")))
    Set StatementText = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Statement Text 2")))
    Set Param2Value = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Parameter 2 Value")))
    Set Source = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Source")))
     
    If Outcome = "Confirmed" Then
       
       ContactType = "Individual"
       PledgePlan = "Continuous"
       PledgeType = "Regular Payment"
       Method = "Recurring"
       ReceiptSummary = "Yes"
       ReceiptReq = "No"
       OneOff = "No"
       CRMPrimaryManager = "jmatchett"
       Param1Name = "In honour of"
       StatementText = "WAYS Phone"
       Param2Value = "NO"
       Source = "WAYS Phone Donations"
       
       If State = "NSW" Or State = "ACT" Then
           SourceCode = "PD40118RE"
       Else
           SourceCode = "PD46018RE"
       End If
       
    End If
    
End Sub
Sub Generate_Static_SG_Fields_Cerebral_Palsy(currentRow As Integer)
    Dim ContactType As Object
    Dim Method As Object
    Dim OneOff As Object
    Dim CRMPrimaryManager As Object
    Dim SourceCode As Object
    Dim Outcome As Object
    Dim State As Object
    
    Set ContactType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Contact Type")))
    Set Method = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Method")))
    Set OneOff = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("One Off?")))
    Set CRMPrimaryManager = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CRM Primary Manager")))
    Set SourceCode = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Source Code")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set State = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("State")))
    
     If Outcome = "Single Gift" Then
       
       ContactType = "Individual"
       Method = "Mail Order"
       OneOff = "Yes"
       CRMPrimaryManager = "jmatchett"
       
       If State = "NSW" Or State = "ACT" Then
           SourceCode = "CD40118RE"
       Else
           SourceCode = "CD46018RE"
       End If
       
    End If
    
End Sub


'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF CHARITY/CAMPAIGN SELECTION----------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|
'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------MISSION AUSTRALIA GENERATION --------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|
'|====================================================================================================================|


Sub Generate_Mission_Reactivations()
 For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Organisation (currentRow)
    Signup_Date (currentRow)
    RegularGiving_Type (currentRow)
    RegularGiving_Date (currentRow)
    OneOfDonation_Date (currentRow)
    Overwrite_Frequency (currentRow)
    Merge_Payment_Type_CC_Type (currentRow)
    Split_Expiry (currentRow)
    Venue_Location_PlaceOfSignUp_FinalCallType (currentRow)
    FinalCallOutcome (currentRow)
    Set_DoNotCall (currentRow)
    Complaint (currentRow)
    Set_Type (currentRow)
 Next currentRow
End Sub

Sub Generate_Mission_Upgrades()
 For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Organisation (currentRow)
    Signup_Date (currentRow)
    RegularGiving_Type (currentRow)
    RegularGiving_Date (currentRow)
    OneOfDonation_Date (currentRow)
    Overwrite_Frequency (currentRow)
    Merge_Payment_Type_CC_Type (currentRow)
    Split_Expiry (currentRow)
    Venue_Location_PlaceOfSignUp_FinalCallType (currentRow)
    FinalCallOutcome (currentRow)
    Set_DoNotCall (currentRow)
    Complaint (currentRow)
    Set_Type (currentRow)
 Next currentRow
End Sub

Sub Organisation(currentRow As Integer)
    Dim Org_Cell As Object
    
    Set Org_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Organisation")))

    Org_Cell = "MA"
End Sub

Sub Signup_Date(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim Signup_Date_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    Dim FinalCallDate As Object

    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Signup_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGivingSignUpDate")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    Set FinalCallDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FinalCallDate")))
    
    FinalCallDate.NumberFormat = "@"
    FinalCallDate = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    
    If Outcome_Cell = "Confirmed" Then
        Signup_Date_Cell.NumberFormat = "@"
        Signup_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    End If
    
       
End Sub

Sub RegularGiving_Type(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim RG_Type_Cell As Object

    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set RG_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGivingType")))
    
    If Outcome_Cell = "Confirmed" Then
        RG_Type_Cell = "Phone"
    End If
End Sub

Sub RegularGiving_Date(currentRow As Integer)
 Dim Outcome_Cell As Object
    Dim StartMonth_Cell As Object
    Dim Debit_Day_Cell As Object
    Dim RegularGivingDate_Cell As Object
    Dim Day As Object
    Dim numMatch As Object
    
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set StartMonth_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    Set Debit_Day_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set RegularGivingDate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGivingDate")))
    Set Day = Debit_Day_Cell
    Set numMatch = CreateObject("vbscript.regexp")
    
    If Outcome_Cell = "Confirmed" Then
    
        numMatch.Pattern = "[0-9]+"
        currentYear = CStr(Year(Now()))
        longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
        If Len(Debit_Day_Cell) > 0 Then
            For Each mon In longMonths
                If mon = StartMonth_Cell Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
            Next mon
              Set num = numMatch.Execute(Day)
            For Each n In num
                Set Day = n
            Next n
            RegularGivingDate_Cell.NumberFormat = "@"
            RegularGivingDate_Cell = Day & "/" & CStr(index) & "/" & currentYear
            RegularGivingDate_Cell = Format(RegularGivingDate_Cell, "dd/mm/yyyy")
            
        Else
            RegularGivingDate_Cell.Interior.ColorIndex = 6
        End If
    
    End If
End Sub

Sub OneOfDonation_Date(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim OneOffDonationDate_Cell As Object
    Dim OutcomeUpdate_Cell As Object

    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set OneOffDonationDate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OneOffDonationDate")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    If Outcome_Cell = "Single Gift" Then
        OneOffDonationDate_Cell.NumberFormat = "@"
        OneOffDonationDate_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    End If
End Sub

Sub Overwrite_Frequency(currentRow As Integer)
    Dim Frequency As Object
    
    Set Frequency = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGivingFrequency")))
    If Len(Frequency) > 0 Then
        Frequency = "MONTHLY"
    End If
End Sub

Sub Merge_Payment_Type_CC_Type(currentRow As Integer)
    Dim Payment As Object
    Dim CardType As Object
    
    Set Payment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
    Set CardType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
    
    If Len(CardType) > 0 Then
        Payment = CardType
    End If
    If Payment = "Direct Debit" Then
        Payment = "Bank Direct Debit"
    End If
End Sub

Sub Split_Expiry(currentRow As Integer)
    Dim Outcome As Object
    Dim Expiry As Object
    Dim Month As Object
    Dim Year As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Expiry = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
    Set Month = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CCExpiryMonth")))
    Set Year = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CCExpiryYear")))
    
    If Outcome = "Confirmed" Or Outcome = "Single Gift" Then
        pos = InStr(Expiry, "/")
        If pos <> 0 Then
            exMonth = Split(Expiry, "/")(0)
            exYear = Split(Expiry, "/")(1)
            Month = exMonth
            Year = exYear
        Else
            Month.Interior.ColorIndex = 6
            Year.Interior.ColorIndex = 6
        End If
    End If

    
End Sub

Sub Venue_Location_PlaceOfSignUp_FinalCallType(currentRow As Integer)
    Dim Place As Object
    Dim Venue As Object
    Dim CallType As Object
    
    Set Place = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PlaceOfSignUp")))
    Set Venue = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("VenueLocation")))
    Set CallType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FinalCallType")))
    
    Place = "Phone"
    Venue = "SYDNEY"
    If Initiative_Name = "Reactivations" Then
        CallType = "Reactivation"
    Else
        CallType = "Upgrades"
    End If

End Sub

Sub FinalCallOutcome(currentRow As Integer)
    Dim Outcome As Object
    Dim FinalOutcome As Object

    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set FinalOutcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FinalCallOutcome")))
    
    If Outcome = "Confirmed" Then
        FinalOutcome = "Successful"
    
    ElseIf Outcome = "Single Gift" Then
    
        FinalOutcome = "RG - Refused but donated one off"
    
    ElseIf Outcome = "Not Interested" Or Outcome = "Already a Supporter" Or Outcome = "Do Not Call" Or Outcome = "Instant Refusal" Or Outcome = "No Survey" Then
        FinalOutcome = "Refused"
    
    ElseIf Outcome = "Wrong Number" Or Outcome = "Disconnected Number" Then
        FinalOutcome = "Wrong Phone Numbers"
    
    ElseIf Outcome = "Uncontactable" Or Outcome = "Completed" Or Outcome = "Max Attempts" Then
        FinalOutcome = "No Answer"
    
    End If
End Sub

Sub Set_DoNotCall(currentRow As Integer)
    Dim Outcome As Object
    Dim NoPhone As Object
    Dim NoTelemarket As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set NoPhone = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
    Set NoTelemarket = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DoNotTelemarket")))
     
    If Outcome = "Do Not Call" Then
        NoPhone = 1
        NoTelemarket = 1
    End If
End Sub

Sub Complaint(currentRow As Integer)
    Dim Outcome As Object
    Dim FurtherNotes As Object

    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set FurtherNotes = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Further Notes")))
    
    If Outcome = "Complaint" Then
    
        FurtherNotes = Outcome
    
    End If

End Sub

Sub Set_Type(currentRow As Integer)
    Dim Type_Cell As Object
    Dim Segment As Object
    
    Set Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Type")))
    Set Segment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("WAYS Segment_1")))
    
    Type_Cell = Segment
    
    
    
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF MISSION AUSTRALIA ------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|
'|====================================================================================================================|



Sub Generate_Upgrade_Amount_WaterAid()
    Dim RG_Cell As Object
    Dim Prev_RG_Cell As Object
    Dim Upgrade_Cell As Object
    Dim Report_Upgrade_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
        Set Prev_RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Current Gift Amount")))
        Set Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Upgrade Amount")))
        Set Report_Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("UG Amount")))
        If Trim(Len(RG_Cell)) <> 0 Then
            newAmount = CInt(RG_Cell)
            oldAmount = CInt(Prev_RG_Cell)
            upgradeAmount = newAmount - oldAmount
            Upgrade_Cell = CStr(upgradeAmount)
            Upgrade_Cell = Format(Upgrade_Cell, "Currency")
            Report_Upgrade_Cell = Upgrade_Cell
            Report_Upgrade_Cell = Format(Report_Upgrade_Cell, "Currency")
            If upgradeAmount < 0 Then
                Upgrade_Cell.Interior.ColorIndex = 6
            End If
        End If
    Next currentRow
End Sub
Sub Generate_Call_Date_WaterAid()
    Dim CallDate_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set CallDate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Last Call Date")))
        Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
        
        CallDate_Cell.NumberFormat = "@"
        CallDate_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    Next currentRow
End Sub


'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------STC SPECIFIC CLEANING----------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH STC SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING       |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

Sub Generate_STC_Reacts()
    Set_Card_Type_STC
    Move_Payment_Details_If_SG_STC
    Generate_CallDate ("Call Date")
    Set_Secondary_Outcome_Desc_If_Deceased_STC
    Convert_No_Phone_To_Y_STC
    Set_RecordingId_STC
    Set_STC_React_Outcome
    Fix_Dates_With_Slash_STC
End Sub

Sub Fix_Dates_With_Slash_STC()
    Dim DOB As Object
    Dim GiftDate As Object
    Dim TransactionDate As Object
    Dim LastDate As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    
        Set DOB = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnBio_Birth_date")))
        Set GiftDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnLGf_1_Date")))
        Set TransactionDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnLGf_1_Nxt_transaction_dat")))
        Set LastDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Last Transaction Date")))
        DOB.NumberFormat = "@"
        DOB = Replace(DOB, ".", "/")
        GiftDate.NumberFormat = "@"
        GiftDate = Replace(GiftDate, ".", "/")
        TransactionDate.NumberFormat = "@"
        TransactionDate = Replace(TransactionDate, ".", "/")
        LastDate.NumberFormat = "@"
        LastDate = Replace(LastDate, ".", "/")
    Next currentRow

End Sub

Sub Set_STC_React_Outcome()
    Dim Outcome As Object
    Dim PrimaryCall As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set PrimaryCall = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Primary Call Outcome Description")))
        
        If Outcome = "Confirmed" Then
            PrimaryCall = "Confirmed"
        ElseIf Outcome = "Single Gift" Then
            PrimaryCall = "Donation"
        ElseIf Outcome = "Disconnected" Or Outcome = "Disconnected Number" Then
            PrimaryCall = "(11) Phone disconnected"
        ElseIf Outcome = "Deceased" Then
            PrimaryCall = "Removed by Request"
        ElseIf Outcome = "Uncontactable" Or Outcome = "Max Attempts" Or Outcome = "Completed" Then
            PrimaryCall = "Not Available"
        ElseIf Outcome = "Wrong Number" Then
            PrimaryCall = "(12) Incorrect Phone Number"
        Else
            PrimaryCall = "Negative"
        End If
    Next currentRow
End Sub

Sub Set_RecordingId_STC()
    
    Dim Callid As Object
    Dim Recording As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        
        Set Callid = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("callid")))
        Set Recording = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RecordingFileNumber")))
        
        Recording = Callid
        
        
    Next currentRow

End Sub

Sub Convert_No_Phone_To_Y_STC()
    Dim No_Phone As Object
     
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set No_Phone = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
        
        
        If Len(Trim(No_Phone)) <> 0 Then
            No_Phone = "Y"
        End If
    Next currentRow
End Sub

Sub Set_Secondary_Outcome_Desc_If_Deceased_STC()
    Dim Outcome As Object
    Dim SecondaryOutcome As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set SecondaryOutcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Secondary Call Outcome Description")))
        
        If Outcome = "Deceased" Then
            SecondaryOutcome = "Deceased"
        End If
    Next currentRow
End Sub

Sub Move_Payment_Details_If_SG_STC()
    Dim Outcome As Object
    Dim CardType As Object
    Dim CardNum As Object
    Dim CardName As Object
    Dim Expiry As Object
    
    Dim SGCardType As Object
    Dim SGCardNum As Object
    Dim SGCardName As Object
    Dim SGExpiry As Object
 
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set CardType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
        Set CardNum = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
        Set CardName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Name on Card")))
        Set Expiry = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
        
        Set SGCardType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OneOff_Credit_Type")))
        Set SGCardNum = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OneOff_Credit_Card_Number")))
        Set SGCardName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OneOff_Cardholder_name")))
        Set SGExpiry = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OneOff_Expiry")))
        
        If Outcome = "Single Gift" Then
            SGCardType = CardType
            SGCardNum = CardNum
            SGCardName = CardName
            SGExpiry = Expiry
            
            CardType = ""
            CardNum = ""
            CardName = ""
            Expiry = ""
        End If
        
    Next currentRow
End Sub

Sub Set_Card_Type_STC()
    Dim CC_Type_Cell As Object
    Dim Start_Month_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set CC_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
        If CC_Type_Cell = "Visa" Then
            CC_Type_Cell = "VISA"
        ElseIf CC_Type_Cell = "Amex" Then
            CC_Type_Cell = "AMEX"
        ElseIf CC_Type_Cell = "Diners" Then
            CC_Type_Cell = "DINERS"
        ElseIf CC_Type_Cell = "Mastercard" Then
            CC_Type_Cell = "MCARD"
        End If
        
    Next currentRow

End Sub

Sub Generate_STC()
 For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Generate_Supplier (currentRow)
    Generate_Call_Date (currentRow)
    Generate_InputTime (currentRow)
    Generate_Call_Outcome (currentRow)
    Generate_DonorSource (currentRow)
    Highlight_Financial_Details_For_NonFinancial_Outcomes (currentRow)
    
    'Fill_Phone_If_Doesnt_Exist (currentRow)
    Clear_DOB (currentRow)
    Clear_Email_STC (currentRow)
    Copy_SG_Values (currentRow)
    Generate_Pensioner (currentRow)
    'Generate_Inbound_Number (currentRow)
    Set_Y_On_Cols (currentRow)
    Generate_Child_Sponsorship (currentRow)
    Set_RG_Product_Code (currentRow)
    Highlight_Feedback_Comments (currentRow)
    Format_Start_Month (currentRow)
    'Format_All_Phones_STC (currentRow)
    

Next currentRow
    Convert_Expiry_to_two_digits
    Fix_Home_Phone_And_Mobile_Phone
    Populate_Title_And_Gender
    Change_Start_Month_Header
End Sub
Sub Change_Start_Month_Header()
    Dim Start As Object
    
    Set Start = Worksheets(Initiative_Name).Cells(1, (Column_Number("Start Month")))
    
    Start = "FirstDebitDate"
     
End Sub

Sub Format_Start_Month(currentRow As Integer)
    Dim Outcome As Object
    Dim StartMonth As Object
    Dim DebitDay As Object
    Dim numMatch As Object
    Dim Day As Object
    
    Set numMatch = CreateObject("vbscript.regexp")
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set DebitDay = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set StartMonth = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    
    If Outcome = "Confirmed" Or Outcome = "Confirmed No Child" Then
        If Len(Trim(StartMonth)) <> 0 And Len(Trim(DebitDay)) Then
      
        numMatch.Pattern = "[0-9]+"
        currentYear = CStr(Year(Now()))
        longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
        For Each mon In longMonths
            If mon = StartMonth Then
                index = Application.Match(mon, longMonths, False) - 1
            End If
        Next mon
          Set num = numMatch.Execute(DebitDay)
        For Each n In num
            Set DebitDay = n
        Next n
        StartMonth.NumberFormat = "@"
        StartMonth = DebitDay & "/" & CStr(index) & "/" & currentYear
        StartMonth = Format(StartMonth, "dd/mm/yyyy")
        
        Else
            StartMonth.Interior.ColorIndex = 6
            DebitDay.Interior.ColorIndex = 6
        End If
    End If

End Sub

'Add Ways Phone to all records
Sub Generate_Supplier(currentRow As Integer)
    Dim Supplier_Cell As Object
    
    Set Supplier_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Supplier")))
    
    Supplier_Cell = "Ways Phone"
    
End Sub
'format outcomeUpdateDateTime Cell to be dd/mm/yyyy
Sub Generate_Call_Date(currentRow As Integer)
    Dim CallDate_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    
    Set CallDate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CallDate")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    CallDate_Cell.NumberFormat = "@"
    CallDate_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
End Sub
'format OutcomeUpdateDateTime to be HH:MM
Sub Generate_InputTime(currentRow As Integer)
    Dim InputTime As Object
    Dim OutcomeUpdate_Cell As Object
    
    Set InputTime = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("InputTime")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    InputTime.NumberFormat = "@"
    InputTime = Format(OutcomeUpdate_Cell, "hh:mm")
End Sub

Sub Fill_Phone_If_Doesnt_Exist(currentRow As Integer)
    Dim CallerId As Object
    Dim Home As Object
    Dim Mobile As Object
    
    Set CallerId = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CallerId")))
    Set Home = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("HomePhone")))
    Set Mobile = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("MobilePhone")))
    
    If Trim(Len(Home)) = 0 Then
        If Trim(Len(CallerId)) <> 0 Then
            If Mid(CallerId, 3, 1) <> "4" Then
            
                Home = Format_Phone_STC(CallerId)
                
            End If
         End If
    End If
    
    If Trim(Len(Mobile)) = 0 Then
        If Trim(Len(CallerId)) <> 0 Then
            If Mid(CallerId, 3, 1) = "4" Then
            
                Mobile = Format_Phone_STC(CallerId)
                
            End If
         End If
    End If
    
End Sub

Sub Format_All_Phones_STC(currentRow As Integer)
    Dim Home As Object
    Dim Mobile As Object
    Dim Work As Object
    Dim NumberDialled As Object
    
    Set Home = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("HomePhone")))
    Set Mobile = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("MobilePhone")))
    Set Work = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Work Phone")))
    
    Set NumberDialled = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("NumberDialled")))
    
    Home = Format_Phone_STC(Home)
    Format_Phone_STC (Mobile)
    Format_Phone_STC (Work)
End Sub

Public Function Format_Phone_STC(CurrentCell As Object)
If Len(Trim(CurrentCell)) <> 0 Then
       
        CurrentCell = Replace(CurrentCell, " ", "")
        If Left(CurrentCell, 2) = "61" Then
            CurrentCell = Replace(CurrentCell, Mid(CurrentCell, 1, 2), "")
        End If
        
        If Left(CurrentCell, 1) <> "4" Then
            If Left(CurrentCell, 1) <> "0" Then
                CurrentCell = "0" & CurrentCell
            ElseIf Mid(CurrentCell, 2, 1) = "4" Then
                CurrentCell = Format(CurrentCell, "0000 000 000")
            End If
            CurrentCell = Format(CurrentCell, "00 0000 0000")
        Else
            If Mid(CurrentCell, 1, 1) <> "0" Then
                CurrentCell = Replace(CurrentCell, " ", "")
                CurrentCell = "0" & CurrentCell
                CurrentCell = Format(CurrentCell, "0000 000 000")
            Else
                CurrentCell = Replace(CurrentCell, " ", "")
                CurrentCell = Format(CurrentCell, "0000 000 000")
            End If
        End If
    End If
    Format_Phone_STC = CurrentCell
End Function

Sub Clear_DOB(currentRow As Integer)
    Dim DOB As Object
    
    Set DOB = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DOB")))
    
    If DOB = "xx/xx/xxxx" Or DOB = "11/11/1900" Or DOB = "11/11/1990" Then
        DOB = ""
    End If
End Sub

Sub Clear_Email_STC(currentRow As Integer)
    Dim Email As Object

    Set Email = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Email")))
    
    If Email = "noemail@gmail.com" Or Email = "noemail@gmail" Then
        Email = ""
    ElseIf InStr(Email, "noemail") <> 0 Then
        Email.Interior.ColorIndex = 6
    End If
End Sub

Sub Copy_SG_Values(currentRow As Integer)
    Dim SG As Object
    Dim RG As Object
    Dim Outcome As Object
    
    
    Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DonationAmount")))
    Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG Amount")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
     
    If Outcome = "Single Gift" Then
        If Len(Trim(RG)) = 0 And Len(Trim(SG)) <> 0 Then
            RG = SG
            RG = Format(RG, "Currency")
        ElseIf Len(Trim(RG)) <> 0 And Len(Trim(SG)) <> 0 Then
            RG.Interior.ColorIndex = 6
            SG.Interior.ColorIndex = 6
        ElseIf Len(Trim(RG)) = 0 And Len(Trim(SG)) = 0 Then
            RG.Interior.ColorIndex = 6
            SG.Interior.ColorIndex = 6
        End If
    ElseIf Len(Trim(SG)) <> 0 Then
        RG.Interior.ColorIndex = 6
        SG.Interior.ColorIndex = 6
    End If
End Sub

Sub Generate_Pensioner(currentRow As Integer)
    Dim Occupation As Object
    Dim Pensioner As Object
    
    Set Occupation = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Occupation")))
    Set Pensioner = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Pensioner")))
    
    If InStr(Occupation, "pension") <> 0 Or InStr(Occupation, "Pension") <> 0 Then
        Pensioner = "Y"
    End If
End Sub

Sub Generate_Inbound_Number(currentRow As Integer)
    Dim NumberDialled As Object
    Dim InboundNumber As Object
    
    Set NumberDialled = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("NumberDialled")))
    Set InboundNumber = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("InboundPhoneNumber")))
    InboundNumber.NumberFormat = "@"
    If Len(Trim(InboundNumber)) <> 0 Then
        Start = Mid(InboundNumber, 1, 3)
        If Start = "611" Then
            InboundNumber = Mid(InboundNumber, 3, Len(InboundNumber))
        ElseIf Mid(InboundNumber, 1, 1) = "6" Then
            InboundNumber.Interior.ColorIndex = 6
        End If
    Else
        InboundNumber = NumberDialled
        Start = Mid(InboundNumber, 1, 3)
        If Start = "611" Then
            InboundNumber = Mid(InboundNumber, 3, Len(InboundNumber))
        End If
    End If
End Sub
Sub Set_Y_On_Cols(currentRow As Integer)
    Dim CCNumCheck As Object
    Dim BSBCCheck As Object
    Dim CCAccountName As Object
    Dim PhoneCheck As Object
    Dim Payment As Object
    
    Set CCNumCheck = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CCNumberChecked")))
    Set BSBCCheck = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BSBChecked")))
    Set CCAccountName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CCAccountName")))
    Set PhoneCheck = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneVerification")))
    Set Payment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))

    PhoneCheck = "Y"
    If Payment = "Credit Card" Then
        CCNumCheck = "Y"
        CCAccountName = "Y"
    ElseIf Payment = "Direct Debit" Then
        BSBCCheck = "Y"
    End If
    
End Sub

Sub Generate_Child_Sponsorship(currentRow As Integer)
    Dim Outcome As Object
    Dim CSPON As Object
    Dim ImmediateGift As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set CSPON = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Reference Number for CSpon")))
    Set ImmediateGift = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DonorAwareofImmediateGift")))
    
    If Len(Trim(CSPON)) <> 0 Then
        ImmediateGift = "Y"
    Else
        If Outcome = "Confirmed" Then
        Outcome.Interior.ColorIndex = 6
        End If
    End If
    
    
End Sub

'if statements to map ways outcomes to STC outcomes
Sub Generate_Call_Outcome(currentRow As Integer)
    Dim STC_Outcome As Object
    Dim Outcome As Object
    Dim Address As Object
    Dim Email As Object
    
    Set STC_Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CallOutcome")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Address = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Address")))
    Set Email = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Email")))
    Set CSPON = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Reference Number for CSpon")))
    
    If Outcome = "Already handled" Then
        STC_Outcome = "Already Handled"
    ElseIf Outcome = "Prank Call" Then
        STC_Outcome = "Prank Call"
    ElseIf Outcome = "Supporter Care" Then
        STC_Outcome = "Supporter Care"
    ElseIf InStr(Outcome, "Promised RG") <> 0 And InStr(Outcome, "no form") <> 0 Then
        STC_Outcome = "Promised RG - no form"
    ElseIf InStr(Outcome, "Promised SG") <> 0 And InStr(Outcome, "no form") <> 0 Then
        STC_Outcome = "Promised SG - no form"
    ElseIf Outcome = "Single Gift" Then
        STC_Outcome = "One Off Gift"
    ElseIf Outcome = "Not Interested" Then
        STC_Outcome = "Not Interested"
        
    ElseIf Outcome = "Promised RG" Then
        If Len(Trim(Address)) <> 0 Then
            STC_Outcome = "Pledged Regular Gift Mail"
        ElseIf Len(Trim(Email)) <> 0 Then
            STC_Outcome = "Pledged Regular Gift Email"
        Else
            STC_Outcome.Interior.ColorIndex = 6
        End If
    
    ElseIf Outcome = "Promised SG" Then
        If Len(Trim(Address)) <> 0 Then
            STC_Outcome = "Pledged Cash Donation Mail"
        ElseIf Len(Trim(Email)) <> 0 Then
            STC_Outcome = "Pledged Cash Donation Email"
        Else
            STC_Outcome.Interior.ColorIndex = 6
        End If
    ElseIf Outcome = "Feedback Only" Or Outcome = "Enquiry Only" Then
        STC_Outcome = "Feedback"
    ElseIf Outcome = "Confirmed No Child" Then
        STC_Outcome = "Regular Gift"
    ElseIf Outcome = "Confirmed" Then
        If Len(Trim(CSPON)) <> 0 Then
            STC_Outcome = "CSPON - Processed Online"
        Else
            STC_Outcome = "Regular Gift"
            STC_Outcome.Interior.ColorIndex = 6
        End If
    ElseIf Outcome = "Completed" Or Outcome = "Max Attempts" Then
        STC_Outcome = "Prank Call"
    End If
    

End Sub

Sub Set_RG_Product_Code(currentRow As Integer)
    Dim Outcome As Object
    Dim Product_Code As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Product_Code = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RGProductCode")))
    
    If Outcome = "Confirmed No Child" Then
        Product_Code = "Child in Crisis"
    End If
    
End Sub

Sub Highlight_Feedback_Comments(currentRow As Integer)
    Dim Comments As Object
    Dim Outcome As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Comments = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Wrap Notes")))
    If (Outcome = "Enquiry Only" Or Outcome = "Feedback Only") And Len(Trim(Comments)) > 0 Then
        Comments.Interior.ColorIndex = 6
    End If
    
End Sub

'Set Donor Source to be TV for all records
Sub Generate_DonorSource(currentRow As Integer)
    Dim DonorSource As Object
    
    Set DonorSource = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DonorSource")))
    
    DonorSource = "TV"
End Sub
'Set Y for all records for Phone verification
Sub Highlight_Financial_Details_For_NonFinancial_Outcomes(currentRow As Integer)
    Dim Outcome As Object
    Dim Payment As Object
    Dim Start As Object
    Dim Debit As Object
    Dim CCNum As Object
    Dim CCType As Object
    Dim CCName As Object
    Dim CCExp As Object
    Dim DDName As Object
    Dim DDNum As Object
    Dim DDBankName As Object
    Dim BSB As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Payment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
    Set Start = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    Set Debit = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set CCNum = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
    Set CCType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
    Set CCName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Name on Card")))
    Set CCExp = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
    Set DDBankName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Bank Name")))
    Set DDNum = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Account Number")))
    Set DDName = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Account Holders Name")))
    Set BSB = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BSB")))
    
    If Outcome <> "Confirmed" And Outcome <> "Single Gift" And Outcome <> "Confirmed No Child" Then
        If Len(Trim(Payment)) <> 0 Then
            Payment.Interior.ColorIndex = 6
        End If
        If Len(Trim(Start)) <> 0 Then
            Start.Interior.ColorIndex = 6
        End If
        If Len(Trim(Debit)) <> 0 Then
            Debit.Interior.ColorIndex = 6
        End If
        If Len(Trim(CCNum)) <> 0 Then
            CCNum.Interior.ColorIndex = 6
        End If
        If Len(Trim(CCType)) <> 0 Then
            CCType.Interior.ColorIndex = 6
        End If
        If Len(Trim(CCName)) <> 0 Then
            CCName.Interior.ColorIndex = 6
        End If
        If Len(Trim(CCExp)) <> 0 Then
            CCExp.Interior.ColorIndex = 6
        End If
        If Len(Trim(DDBankName)) <> 0 Then
            DDBankName.Interior.ColorIndex = 6
        End If
        If Len(Trim(DDNum)) <> 0 Then
            DDNum.Interior.ColorIndex = 6
        End If
        If Len(Trim(DDName)) <> 0 Then
            DDName.Interior.ColorIndex = 6
        End If
         If Len(Trim(BSB)) <> 0 Then
            BSB.Interior.ColorIndex = 6
        End If
    End If
    If Outcome = "Single Gift" Then
      If Len(Trim(DDBankName)) <> 0 Then
            DDBankName.Interior.ColorIndex = 6
        End If
        If Len(Trim(DDNum)) <> 0 Then
            DDNum.Interior.ColorIndex = 6
        End If
        If Len(Trim(DDName)) <> 0 Then
            DDName.Interior.ColorIndex = 6
        End If
         If Len(Trim(BSB)) <> 0 Then
            BSB.Interior.ColorIndex = 6
        End If
    End If
    
    
End Sub



'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------MAW SPECIFIC CLEANING----------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH MAW SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING             |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|



Sub Generate_MAW()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Generate_Deceased (currentRow)
        Set_Wrong_Numbers (currentRow)
        Highlight_New_Address_If_Not_Valid (currentRow)
        Join_RG_SG (currentRow)
        Highlight_Incorrect_Report_Amount (currentRow)
        Gen_Start_Date_MAW (currentRow)
        Gen_Call_Date (currentRow)
        Format_As_Number (currentRow)
        Gen_Receipt_Type (currentRow)
        Set_ValidAddress (currentRow)
        Gen_TM_Recording (currentRow)
        Generate_MAW_Outcomes (currentRow)
        Change_No_Phone (currentRow)
    Next currentRow
    Change_No_Phone_Name
End Sub

Sub Change_No_Phone_Name()
    Dim No_Phone_Header As Object
    Set No_Phone_Header = Worksheets(Initiative_Name).Cells(1, (Column_Number("No Phone")))
    No_Phone_Header = "Do Not Contact"
End Sub

Sub Generate_Deceased(currentRow As Integer)
    Add_Y_If_Deceased (currentRow)
    Add_Date_If_Deceased (currentRow)
End Sub
Sub Add_Y_If_Deceased(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim Deceased_Cell As Object

    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Deceased_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("10_New_Deceased")))
    
    If Outcome_Cell = "Deceased" Then
        Deceased_Cell = "Y"
    End If
End Sub
Sub Add_Date_If_Deceased(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim DeceasedDate_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set DeceasedDate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("11_New_Deceased Date")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    If Outcome_Cell = "Deceased" Then
        DeceasedDate_Cell.NumberFormat = "@"
        DeceasedDate_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    End If
End Sub

Sub Gen_Receipt_Type(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim ReceiptType_Cell As Object
     
     Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
     Set ReceiptType_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Receipt Preference")))
     
     If Outcome_Cell = "Confirmed" Then
        ReceiptType_Cell = "Consolidated"
     Else
        ReceiptType_Cell = "Single Receipt"
     End If
End Sub



Sub Gen_Start_Date_MAW(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim StartMonth_Cell As Object
    Dim Debit_Day_Cell As Object
    Dim ScheduleStart_Cell As Object
    
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set StartMonth_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("62_Start Month")))
    Set Debit_Day_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set ScheduleStart_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Date")))
    
    
    If Outcome_Cell = "Confirmed" Then
    
        'numMatch.Pattern = "[0-9]+"
        currentYear = CStr(Year(Now()))
        longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
   
        
        If Len(Debit_Day_Cell) > 0 Then
            For Each mon In longMonths
                If mon = StartMonth_Cell Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
            Next mon
            
            ScheduleStart_Cell.NumberFormat = "@"
            ScheduleStart_Cell = Debit_Day_Cell & "/" & CStr(index) & "/" & currentYear
            ScheduleStart_Cell = Format(ScheduleStart_Cell, "dd/mm/yyyy")
            
        Else
            ScheduleStart_Cell.Interior.ColorIndex = 6
        End If
    
    End If
End Sub


Sub Gen_TM_Recording(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim TMRecordingID_Cell As Object
    Dim CallId_Cell As Object
    Dim Payment As Object
    
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set CallId_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("callid")))
    Set TMRecordingID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("60_TM Recording ID")))
    Set Payment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
     
    If Outcome_Cell = "Confirmed" And Payment = "Direct Debit" Then
        TMRecordingID_Cell = CallId_Cell
    End If
    If Outcome_Cell = "Single Gift" And Payment = "Direct Debit" Then
        Payment.Interior.ColorIndex = 6
    End If
End Sub

Sub Change_No_Phone(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim No_Phone_Cell As Object

    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set No_Phone_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
    
    If Outcome_Cell = "Do Not Call" Then
        No_Phone_Cell = "Y"
    End If
End Sub

Sub Generate_Static_Values(currentRow As Integer)
    Set_Wrong_Numbers (currentRow)
    Highlight_New_Address_If_Not_Valid (currentRow)
    Join_RG_SG (currentRow)
    Highlight_Incorrect_Report_Amount (currentRow)
    Gen_Start_Date_MAW (currentRow)
    Gen_Call_Date (currentRow)
    Format_As_Number (currentRow)
    Gen_Receipt_Type (currentRow)
    Set_ValidAddress (currentRow)
    Gen_TM_Recording (currentRow)
    Generate_MAW_Outcomes (currentRow)
End Sub
Sub Set_ValidAddress(currentRow As Integer)
    Dim New_Address As Object
    Dim Valid As Object
    
    Set New_Address = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Mailing Street")))
    Set Valid = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_ValidAddress")))
    
    If Len(Trim(New_Address)) > 0 Then
        Valid = "Y"
    End If

End Sub

Sub Format_As_Number(currentRow As Integer)
    Dim Debit As Object
    Dim ExpiryMonth As Object
    Dim ExpiryYear As Object
    
    Set Debit = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set ExpiryMonth = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ExpiryMonth")))
    Set ExpiryYear = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ExpiryYear")))
    
    
    Debit.NumberFormat = "0"
    Debit = Format(Debit, "0")
    ExpiryMonth.NumberFormat = "0"
    ExpiryMonth = Format(ExpiryMonth, "0")
    ExpiryYear.NumberFormat = "0"
    ExpiryYear = Format(ExpiryYear, "0")

End Sub

Sub Highlight_Incorrect_Report_Amount(currentRow As Integer)
    Dim RG As Object
    Dim Outcome As Object
    Dim SG As Object
    
    Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG")))
    
    If Outcome <> "Confirmed" And Len(Trim(RG)) > 0 Then
        RG.Interior.ColorIndex = 6
    End If
    If Outcome <> "Single Gift" And Len(Trim(SG)) > 0 Then
        SG.Interior.ColorIndex = 6
    End If
End Sub

Sub Join_RG_SG(currentRow As Integer)
    Dim Amount As Object
    Dim SG As Object
    Dim RG As Object
    Dim SG_Report As Object
    Dim Outcome As Object
    
    Set Amount = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Amount")))
    Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("61_SG Amount")))
    Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
    Set SG_Report = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG")))

    RG = Amount
    RG = Format(RG, "Currency")
    SG_Report = Format(SG, "Currency")
    
    If (Len(Trim(SG)) > 0) Then
        Amount = Format(SG, "Standard")
    End If
    SG = Format(SG, "Currency")
    SG.NumberFormat = "0.00"
    Amount.NumberFormat = "0.00"
End Sub

Sub Set_Wrong_Numbers(currentRow As Integer)
    Dim Outcome As Object
    Dim Home As Object
    Dim Mobile As Object
    Dim Wrong_Home As Object
    Dim Wrong_Mobile As Object
    
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Home = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Home Phone")))
    Set Mobile = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Mobile")))
    Set Wrong_Home = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("27_Wrong_Home Number")))
    Set Wrong_Mobile = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("30_Wrong_Mobile Number")))
    
    If Outcome = "Wrong Number" Or Outcome = "Disconnected" Or Outcome = "Disconnected Number" Then
    
        Wrong_Home = Home
        Wrong_Mobile = Mobile
    
    End If
End Sub

Sub Highlight_New_Address_If_Not_Valid(currentRow As Integer)
    Dim Valid_Address As Object
    Dim New_Address As Object
    Dim New_PostCode As Object
    Dim New_State As Object
    Dim New_Suburb As Object
    Dim Outcome As Object
    
    Set Valid_Address = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ValidAddress")))
    Set New_Address = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Mailing Street")))
    Set New_PostCode = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Mailing Postal Code")))
    Set New_State = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Mailing State")))
    Set New_Suburb = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("New_Mailing City")))
    Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    If (Valid_Address = "N" And (Outcome = "Confirmed" Or Outcome = "Single Gift" Or Outcome = "Promised RG" Or Outcome = "Promised SG")) Then
        New_PostCode.Interior.ColorIndex = 6
        New_State.Interior.ColorIndex = 6
        New_Suburb.Interior.ColorIndex = 6
        New_Address.Interior.ColorIndex = 6
    Else
        New_PostCode.Interior.ColorIndex = xlNone
        New_State.Interior.ColorIndex = xlNone
        New_Suburb.Interior.ColorIndex = xlNone
        New_Address.Interior.ColorIndex = xlNone
    End If
    
End Sub


Sub Gen_Call_Date(currentRow As Integer)
    Dim Call_Date_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    
    Set Call_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("38_Call Date")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    Call_Date_Cell.NumberFormat = "@"
    Call_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")

End Sub

Sub Generate_MAW_Outcomes(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim Call_Response_Cell As Object
     
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Call_Response_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("37_Response")))
     
    If Outcome_Cell = "Confirmed" Then
        Call_Response_Cell = "Confirmed"
    ElseIf Outcome_Cell = "Single Gift" Then
        Call_Response_Cell = "Donation"
    ElseIf Outcome_Cell = "Already a Supporter" Or Outcome_Cell = "Deceased" Or Outcome_Cell = "Do Not Call" Or Outcome_Cell = "Instant Refusal" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Not Interested" Or Outcome_Cell = "Promised SG" Or Outcome_Cell = "Promised RG" Then
        Call_Response_Cell = "Negative"
    ElseIf Outcome_Cell = "Uncontactable" Or Outcome_Cell = "Max Attempts" Or Outcome_Cell = "Completed" Then
        Call_Response_Cell = "Not Available"
    ElseIf Outcome_Cell = "Disconnected" Or Outcome_Cell = "Wrong Number" Then
        Call_Response_Cell = "Wrong Number"
    Else
        Call_Response_Cell.Interior.ColorIndex = 6
    End If
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------Greenpeace SPECIFIC CLEANING----------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH Greenpeace SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING       |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

Sub Bight_Generation()
    Generate_GP_Outcomes
    Fill_Email_Update
    Format_Expiry_Date
    Generate_CallDate_SignUpDate
End Sub
Sub Greenpeace_Generation()
    Generate_CallDate_SignUpDate
    Generate_recordingRef_Fundraiser
    Replace_CCType_PayMethod
    Generate_GP_Outcomes
    Fill_Email_Update
    Format_Expiry_Date
    Generate_Agency_GP
End Sub
Sub Greenpeace_DRTV_Gen()
    Generate_CallDate_SignUpDate
    Generate_recordingRef_Fundraiser
    Replace_CCType_PayMethod
    Generate_GP_Outcomes
    Format_Expiry_Date
    Generate_Agency_GP
End Sub

Sub Generate_Agency_GP()
Dim List_Cell As Object
Dim Agency As Object

'what we look for in the list name
shortSources = Array("8th", "egentic", "embr", "flagship", "fresh", "kobi", "luna", "marketing", "offers", "quinn", "upside", "zinq", "reach", "drtv", "open")

'what we use to put in the acq source column
LongSources = Array("Ways 8thfloor", "Ways eGENTIC", "Ways EMBR", "Ways flagshipDigital", "Ways freshbackmedia", "WAYS kobi", "WAYS lunaparkmedia", "Ways marketingpunch", "Ways offersNow", "Ways Quinn", "Ways upside", "Ways Zinq", "Ways ReachTel", "Ways DRTV Arctic", "WAYS Opentop")

For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set List_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DataSetName")))
    Set Agency = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Agency")))
    isSource = False
    temp = StrConv(List_Cell, vbLowerCase)
    For Each Source In shortSources
        pos = InStr(temp, Source)
        If pos <> 0 Then
            index = Application.Match(Source, shortSources, False) - 1
            Agency = LongSources(index)
            isSource = True
        End If
    Next Source
    If isSource = False Then
        Agency.Interior.ColorIndex = 6
    End If
Next currentRow
End Sub

Sub Fill_Email_Update()
    Dim EmailUpdate_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set EmailUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("EmailUpdates")))
        EmailUpdate_Cell = "Yes"
    Next currentRow
End Sub

Sub Generate_GP_Outcomes()
    Dim GPOutcome_Cell As Object
    Dim Outcome_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set GPOutcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CallOutcome")))
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        
        If Outcome_Cell = "Confirmed" Then
            GPOutcome_Cell = "YES TO ASK"
        ElseIf Outcome_Cell = "Single Gift" Then
            GPOutcome_Cell = "ONE-OFF DONATION"
        
        ElseIf Outcome_Cell = "Wrong Number" Or Outcome_Cell = "Disconnected" Then
            GPOutcome_Cell = "WRONG NUMBER"
        
        ElseIf Outcome_Cell = "Promised SG" Or Outcome_Cell = "Promised RG" Then
            GPOutcome_Cell = "PROMISED"
        
        ElseIf Outcome_Cell = "Uncontactable" Or Outcome_Cell = "Max Attempts" Or Outcome_Cell = "Completed" Then
            GPOutcome_Cell = "NOT AVAILABLE"
        
        ElseIf Outcome_Cell = "Not Interested" Or Outcome_Cell = "Instant Refusal" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Already a Supporter" Then
            GPOutcome_Cell = "NO TO ASK"
        
        ElseIf Outcome_Cell = "Do Not Call" Then
            GPOutcome_Cell = "DO NOT CALL"
            
        ElseIf Outcome_Cell = "Deceased" Then
            GPOutcome_Cell = "DECEASED"
        Else
            GPOutcome_Cell.Interior.ColorIndex = 6
        End If
    
    Next currentRow
End Sub


Sub Replace_CCType_PayMethod()
    Dim CCType_Cell As Object
    Dim PayMeth_Cell As Object
    Dim BSB_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set CCType_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
        Set BSB_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BSB")))
        Set PayMeth_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        
        If CCType_Cell = "Mastercard" Or CCType_Cell = "Visa" Or CCType_Cell = "Amex" Then
            PayMeth_Cell = "CC"
        End If
        If CCType_Cell = "Mastercard" Then
            CCType_Cell = "M"
        End If
        If CCType_Cell = "Visa" Then
            CCType_Cell = "V"
        End If
        If CCType_Cell = "Amex" Then
            CCType_Cell = "A"
        End If
        If BSB_Cell <> "" And CCType_Cell = "" Then
            PayMeth_Cell = "DD"
        End If
    Next currentRow
End Sub

Sub Generate_CallDate_SignUpDate()
 Dim Call_Date_Cell As Object
 Dim OutcomeUpdate_Cell As Object
 Dim Signup_Date_Cell As Object
 Dim Outcome_Cell As Object
 
 For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Call_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CallDate")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    Set Signup_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SignUpDate")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    Call_Date_Cell.NumberFormat = "@"
    Call_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    If Outcome_Cell = "Confirmed" Then
        Signup_Date_Cell.NumberFormat = "@"
        Signup_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    End If
Next currentRow
End Sub

Sub Generate_recordingRef_Fundraiser()
Dim CallId_Cell As Object
Dim RecordingRef_Cell As Object
Dim AgentName_Cell As Object
Dim Fundraiser_Cell As Object
 
 For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set CallId_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("callid")))
    Set RecordingRef_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RecordingRef")))
    Set AgentName_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AgentName")))
    Set Fundraiser_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Fundraiser")))

    RecordingRef_Cell = CallId_Cell
    Fundraiser_Cell = AgentName_Cell
    
Next currentRow
End Sub


Sub Opportunity_Generation()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Generate_Call_Date (currentRow)
        Generate_DebitDate (currentRow)
    Next currentRow
End Sub


'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------Starlight SPECIFIC CLEANING----------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH Starlight SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING       |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|


Sub Starlight_Generation()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Call_Date (currentRow)
        Generate_SignUpDate (currentRow)
        Generate_DebitDate (currentRow)
    Next currentRow
End Sub

Sub Call_Date(currentRow As Integer)
    Dim Call_Date_Cell As Object
    Dim OutcomeUpdate_Cell As Object

    Set Call_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Call Date")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))

    Call_Date_Cell.NumberFormat = "@"
    Call_Date_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")

End Sub

Sub Generate_SignUpDate(currentRow As Integer)
    Dim Outcome_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    Dim Signup_Cell As Object
    
    Set Signup_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Sign Up Date")))
    Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    If Outcome_Cell = "Confirmed" Then
        Signup_Cell.NumberFormat = "@"
        Signup_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    End If

End Sub

Sub Generate_DebitDate(currentRow As Integer)
    Dim Start_Month_Cell As Object
    Dim Debit_Date_Cell As Object
    Dim Outcome_Cell As Object
    Dim Day As Object
    Dim numMatch As Object
    
    Set Start_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    Set Debit_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Day = Debit_Date_Cell
    Set numMatch = CreateObject("vbscript.regexp")
    
    numMatch.Pattern = "[0-9]+"
    currentYear = CStr(Year(Now()))
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
   
        
        If Len(Debit_Date_Cell) > 0 And (Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child") Then
            For Each mon In longMonths
                If mon = Start_Month_Cell Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
            Next mon
            Set num = numMatch.Execute(Day)
            For Each n In num
                Set Day = n
            Next n
            Debit_Date_Cell.NumberFormat = "@"
            Debit_Date_Cell = Day & "/" & CStr(index) & "/" & currentYear
            Debit_Date_Cell = Format(Debit_Date_Cell, "dd/mm/yyyy")
            
        Else
            Debit_Date_Cell = Null
        End If
    
    
End Sub


'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------WWF SPECIFIC CLEANING----------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH WWF SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING             |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

Sub WWF_Generation()
    Generate_Appeal_Code
    Generate_Donor_Source
    Generate_Final_Result_Date_WWF
    Generate_Final_Result_WWF
    Generate_Debit_Date_WWF
    Copy_SG_to_RG_Column_WWF
    Fix_Card_Type_WWF
    Generate_Source_Group
    Set_No_Phone_As_Y
    Set_CallId_As_Vox_For_DirectDebit
    Set_Negative_Reason
End Sub

Sub WWF_Petition_Generation()
    Generate_Final_Result_Date_WWF
    Generate_Final_Result_WWF
    Generate_Debit_Date_WWF
    Copy_SG_to_RG_Column_WWF
    Fix_Card_Type_WWF
    Set_No_Phone_As_Y
    Set_CallId_As_Vox_For_DirectDebit
    Set_Negative_Reason
   
End Sub

Sub Copy_SG_to_RG_Column_WWF()
    Dim SG As Object
    Dim RG As Object
    Dim Report_RG As Object
    Dim Outcome As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    
        Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG Amount")))
        Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
        Set Report_RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG")))
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        
        Report_RG = Format(RG, "currency")
        
        If Outcome = "Confirmed" Then
            If Len(Trim(Report_RG)) <> 0 Then
                Report_RG.Interior.ColorIndex = 0
            Else
                Report_RG.Interior.ColorIndex = 6
            End If
        End If
        If Len(Trim(RG)) = 0 Then
           
            RG = Format(SG, "currency")
        End If
        
    Next currentRow
End Sub

Sub Set_Acq_Source()
    Dim Acq_Source As Object
    Dim Source_Group As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    
        Set Acq_Source = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Acq Source")))
        Set Source_Group = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Source_group")))
        
        Acq_Source = Source_Group
        
    Next currentRow
End Sub

Sub Set_Negative_Reason()
    Dim Outcome As Object
    Dim Reason As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Reason = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Reason for result")))
        
        If Trim(Len(Reason)) = 0 Then
            If Outcome = "Not Interested" Or Outcome = "Do Not Call" Or Outcome = "Instant Refusal" Then
                Reason = "No Reason Given"
            ElseIf Outcome = "Already a Supporter" Or Outcome = "No Survey" Or Outcome = "Do Not Call" Then
                Reason = "Other Reason"
            End If
        End If
        
    Next currentRow

End Sub

Sub Set_CallId_As_Vox_For_DirectDebit()
    Dim Payment As Object
    Dim Outcome As Object
    Dim Vox As Object
    Dim Callid As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Payment = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Vox = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Vox file number")))
        Set Callid = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("callid")))
        
        If Outcome = "Confirmed" Then
            If Payment = "Direct Debit" Then
                Vox = Callid
            End If
        End If
        
    Next currentRow

End Sub

Sub Set_No_Phone_As_Y()
Dim NoPhone As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set NoPhone = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))
        If (Trim(Len(NoPhone))) > 0 Then
            NoPhone = "Y"
        End If
    Next currentRow
End Sub

Sub Convert_Expiry_to_two_digits()
 Dim Expiry_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Expiry_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
        If Len(Trim(Expiry_Cell)) > 0 Then
            If Len(Trim(Expiry_Cell)) = 7 Then
                exYear = Split(Expiry_Cell, "/")(1)
                exYear = Mid(exYear, 3, 2)
                Expiry_Cell = Split(Expiry_Cell, "/")(0) + "/" + exYear
            Else
                Expiry_Cell.Interior.ColorIndex = 6
            End If
        End If
        
    Next currentRow

End Sub


Sub Generate_Source_Group()
  Dim Source_Group_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Source_Group_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Source_group")))
        
        Source_Group_Cell = "Telemarketing"
    Next currentRow

End Sub


Sub Generate_Appeal_Code()
    Dim Appeal_Code_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Appeal_Code_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AppealCode")))
        If Initiative_Type = "Lead Conversion" Then
            Appeal_Code_Cell = "ONLACQ16-17PS"
        End If
    Next currentRow

End Sub

Sub Generate_Donor_Source()
    Dim Donor_Source_Cell As Object
    Dim Acq_Source_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Donor_Source_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("donor_source")))
        Set Acq_Source_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Acq Source")))
        If Initiative_Type = "Lead Conversion" Then
            Donor_Source_Cell = Acq_Source_Cell
        ElseIf Initiative_Type = "Recycled" Then
            Donor_Source_Cell = "Telemarketing"
        End If
    Next currentRow

End Sub

Sub Generate_Final_Result_Date_WWF()
    Dim OutcomeUpdate_Cell As Object
    Dim Date_Of_Final_Result As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
        Set Date_Of_Final_Result = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Date of final result")))
        Date_Of_Final_Result.NumberFormat = "@"
        Date_Of_Final_Result = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    Next currentRow

End Sub

Sub Generate_Final_Result_WWF()
    Dim Final_Result_Cell As Object
    Dim Outcome_Cell As Object
    finalResultCol = Column_Number("Final result")
    outcomeCol = Column_Number("Outcome")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Final_Result_Cell = Worksheets(Initiative_Name).Cells(currentRow, finalResultCol)
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, outcomeCol)
        If Outcome_Cell = "Confirmed" Then
            Final_Result_Cell = "Confirmed"
        ElseIf Outcome_Cell = "Not Interested" Or Outcome_Cell = "Already a Supporter" Or _
        Outcome_Cell = "Instant Refusal" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Do Not Call" Then
            Final_Result_Cell = "Negative"
        ElseIf Outcome_Cell = "Uncontactable" Or Outcome_Cell = "Completed" Or Outcome_Cell = "Max Attempts" Then
            Final_Result_Cell = "Not Available"
        ElseIf Outcome_Cell = "Wrong Number" Or Outcome_Cell = "Disconnected" Or Outcome_Cell = "Disconnected Number" Then
            Final_Result_Cell = "Wrong Number"
        ElseIf Outcome_Cell = "Deceased" Then
            Final_Result_Cell = "Deceased"
        ElseIf Outcome_Cell = "Single Gift" Then
            Final_Result_Cell = "Donation"
        Else
            Final_Result_Cell.Interior.ColorIndex = 6
        End If
    Next currentRow
End Sub

Sub Generate_Debit_Date_WWF()
    Dim Debit_Cell As Object
    Dim Start_Month_Cell As Object
    Dim Day As Object
    Dim DateString As Object
    Dim numMatch As Object
    Dim Outcome_Cell As Object
    
    
    Set numMatch = CreateObject("vbscript.regexp")
    numMatch.Pattern = "[0-9]+"
    currentYear = CStr(Year(Now()))
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Debit_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Starting Date")))
        Set Start_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Day = Debit_Cell
        If Len(Debit_Cell) > 0 And (Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child") Then
            For Each mon In longMonths
                If mon = Start_Month_Cell Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
            Next mon
            Set num = numMatch.Execute(Day)
            For Each n In num
                Set Day = n
            Next n
            Debit_Cell.NumberFormat = "@"
            Debit_Cell = Day & "/" & CStr(index) & "/" & currentYear
            
        Else
            Debit_Cell = Null
        End If
    Next currentRow

End Sub
Sub Fix_Card_Type_WWF()
    Dim CC_Type_Cell As Object
    Dim Start_Month_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set CC_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Type")))
        If CC_Type_Cell = "Visa" Then
            CC_Type_Cell = "VISA"
        ElseIf CC_Type_Cell = "Amex" Then
            CC_Type_Cell = "AMEX"
        ElseIf CC_Type_Cell = "Diners" Then
            CC_Type_Cell = "DINERS"
        End If
        
    Next currentRow
End Sub





'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------THE SMITH FAMILY SPECIFIC CLEANING---------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'| THIS SECTION OF THE MACRO IS CONCERNED WITH THE SMITH FAMILY SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING   |
'| START MONTH AND IMPORT RESULT.                                                                                     |
'|____________________________________________________________________________________________________________________|


Sub Generate_DonationType_SmithFamily()
    Dim RG As Object
    Dim Outcome As Object
    Dim DonationType As Object
    Dim DonationType2 As Object
    Dim SG As Object
    Dim NoOfChildren As Object
    Dim NoOfChildren2 As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set RG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Amount")))
        Set Outcome = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set DonationType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Donation type")))
        Set DonationType2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Donation type2")))
        Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG Amount")))
        Set NoOfChildren = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No of school students")))
        Set NoOfChildren2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No of school students2")))
        
        NoOfChildren = Format(NoOfChildren, "#.00")
        
        If Outcome = "Confirmed" Then
            If RG Mod 48 = 0 Then
                DonationType = "Sponsorship"
                
            Else
                NoOfChildren.Interior.ColorIndex = 6
                If (Len(Trim(DonationType))) = 0 Then
                    DonationType = "MonthlyDonation"
                    Outcome.Interior.ColorIndex = 6
                Else
                    DonationType.Interior.ColorIndex = 6
                    Outcome.Interior.ColorIndex = 6
                End If
                
            End If
        
        ElseIf Outcome = "Confirmed No Child" Then
             If Len(Trim(NoOfChildren)) <> 0 Then
                NoOfChildren.Interior.ColorIndex = 6
             End If
             If RG Mod 48 = 0 Then
                If (Len(Trim(DonationType))) = 0 Then
                    DonationType = "Sponsorship"
                    Outcome.Interior.ColorIndex = 6
                    
                Else
                    DonationType.Interior.ColorIndex = 6
                    Outcome.Interior.ColorIndex = 6
                End If
            Else
                DonationType = "MonthlyDonation"
            End If
        ElseIf Outcome = "Single Gift" Then
            DonationType = "OneOffDonation"
        Else
            If Len(Trim(RG)) <> 0 Then
                RG.Interior.ColorIndex = 6
            End If
            If Len(Trim(SG)) <> 0 Then
                SG.Interior.ColorIndex = 6
            End If
            If Len(Trim(DonationType)) <> 0 Then
                DonationType.Interior.ColorIndex = 6
            End If
        End If
        DonationType2 = DonationType
        NoOfChildren2 = NoOfChildren
    Next currentRow
End Sub

Sub Generate_ArrearsAmount_TSF_Rejects()
    Dim SG As Object
    Dim Arrears  As Object
  
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    
        Set SG = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("SG Amount")))
        Set Arrears = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Arrears Amount*")))
      
        If Len(Trim(SG)) > 0 Then
            Arrears = SG
            Arreats = Format(Arrears, "Number")
        End If
    Next currentRow

End Sub

Sub Copy_ConsumerId_To_Supporter_ID_TSF_Rejects()
  Dim CustKey As Object
  Dim SupporterID  As Object

  For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
  
      Set CustKey = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Consumer ID")))
      Set SupporterID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Supporter ID")))
    
      SupporterID = CustKey
      
  Next currentRow

End Sub


Sub Generate_Country()
    Dim Country_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Country_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Country")))
        Country_Cell = "AUSTRALIA"
    Next currentRow
End Sub

Sub Generate_Agency()
    Dim Agency_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Agency_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Agency Name*")))
        Agency_Cell = "WAYSPHONE"
    Next currentRow
End Sub

Sub Generate_Signup_Date()
    Dim OutcomeUpdate_Cell As Object
    Dim Signup_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
        Set Signup_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Sign-Up Date*")))
        Signup_Cell.NumberFormat = "@"
        Signup_Cell = Format(OutcomeUpdate_Cell, "dd/mm/yyyy")
    Next currentRow

End Sub

Sub Generate_Debit_Date()
    Dim Debit_Cell As Object
    Dim Start_Month_Cell As Object
    Dim Day As Object
    Dim DateString As Object
    Dim numMatch As Object
    Dim Outcome_Cell As Object
    
    
    Set numMatch = CreateObject("vbscript.regexp")
    numMatch.Pattern = "[0-9]+"
    currentYear = CStr(Year(Now()))
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Debit_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Date")))
        Set Start_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Day = Debit_Cell
        If Len(Debit_Cell) > 0 And (Outcome_Cell = "Confirmed" Or Outcome_Cell = "Confirmed No Child") Then
            For Each mon In longMonths
                If mon = Start_Month_Cell Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
            Next mon
            Set num = numMatch.Execute(Day)
            For Each n In num
                Set Day = n
            Next n
            Debit_Cell.NumberFormat = "@"
            If index < 10 Then
                index = "0" + CStr(index)
            Else
                index = CStr(index)
            End If
            Debit_Cell = Day & "/" & index & "/" & currentYear
            
        Else
            Debit_Cell = Null
        End If
    Next currentRow

End Sub

Sub Generate_Payment_method()
    Dim Payment_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Payment_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        If Payment_Cell = "Direct Debit" Then
            Payment_Cell = "Bankaccount"
        ElseIf Payment_Cell = "Credit Card" Then
            Payment_Cell = "Creditcard"
        End If
    Next currentRow
End Sub




'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------SALVOS SPECIFIC CLEANING----------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH SALVOS SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING             |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

 Sub Check_Token_Card_Salvos()
    Dim Token_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Token_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
        If Len(Token_Cell) <> 0 Then
            pos = InStr(Token_Cell, "0510")
            If pos = 0 Then
                Token_Cell.Interior.ColorIndex = 6
            End If
        End If
    Next currentRow
        
    
 
 End Sub

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------PETER MAC SPECIFIC CLEANING----------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH PETER MAC SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING       |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

Sub Generate_Peter_Mac()
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Generate_File_Ids (currentRow)
        Generate_KeyInd (currentRow)
        Generate_PrimAddID (currentRow)
        Generate_PrimSalID (currentRow)
        Generate_AddrImpID (currentRow)
        Generate_PhoneImpID1 (currentRow)
        Generate_PhoneImpID2 (currentRow)
        Generate_PhoneImpID3 (currentRow)
        Generate_Outcome_cells (currentRow)
        Generate_GFImpID (currentRow)
        Generate_GFInsStartDate (currentRow)
        Generate_State (currentRow)
        Generate_PM_Outcomes (currentRow)
        Fix_No_Phone (currentRow)
        Generate_Unknown_Gender_Title (currentRow)
    Next currentRow
    Write_PM_ID_NUM
End Sub
Sub Generate_Unknown_Gender_Title(currentRow As Integer)
    Dim Title_Cell As Object
    Dim Gender_Cell As Object
    
    Set Title_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Title")))
    Set Gender_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Gender")))
    
    If Len(Trim(Gender_Cell)) = 0 Then
        Gender_Cell = "Unknown"
        Gender_Cell.Interior.ColorIndex = 6
    End If
    If Len(Trim(Title_Cell)) = 0 Then
        Title_Cell = "M/s"
        Title_Cell.Interior.ColorIndex = 6
    End If

End Sub


Sub Write_PM_ID_NUM()
    Workbooks("WAYS Data Cleaning Macro").Sheets("PM_ID_NUM").Cells(1, 1) = PETER_MAC_FILE_ID
End Sub

Sub Generate_File_Ids(currentRow As Integer)
    Dim Cons_ID As Object
    Dim Import_ID As Object
    Dim GFReference As Object
    Dim Client_ID As Object
    
    
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set Import_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ImportID")))
    Set GFReference = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFReference")))
    Set Client_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Client_ID")))
    
    Input_ID_Length = Len(CStr(PETER_MAC_FILE_ID))
    zeroCount = 4 - (Input_ID_Length)
    zeroes = ""
    For Z = 0 To zeroCount Step 1
        zeroes = zeroes & "0"
    Next Z
    
    Cons_ID = "WT" & zeroes & CStr(PETER_MAC_FILE_ID)
    PETER_MAC_FILE_ID = PETER_MAC_FILE_ID + 1
    Import_ID = Cons_ID
    GFReference = Cons_ID
    Client_ID = Cons_ID

End Sub

Sub Fix_No_Phone(currentRow As Integer)
Dim No_Phone_Cell As Object

 Set No_Phone_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("No Phone")))

If Len(Trim(No_Phone_Cell)) > 0 Then
    No_Phone_Cell = "Yes"
End If

End Sub

Sub Generate_PM_Outcomes(currentRow As Integer)
    Dim Primary_Cell As Object
    Dim Secondary_Cell As Object
    Dim Payment_Cell As Object
    Dim Outcome_Cell As Object
    
    Set Primary_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Primary Call Outcome")))
    Set Secondary_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Secondary Call Outcome")))
    Set Payment_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    
    If Outcome_Cell = "Confirmed" Then
        If Payment_Cell = "Credit Card" Then
            Primary_Cell = "Contact Made"
            Secondary_Cell = "Monthly Amt & CC details given over phone"
        ElseIf Payment_Cell = "Direct Debit" Then
            Primary_Cell = "Contact Made"
            Secondary_Cell = "Monthly Amt & DD Acct details given over Phone"
        Else
            Primary_Cell = "Contact Made"
            Secondary_Cell.Interior.ColorIndex = 6
        End If
        
    ElseIf Outcome_Cell = "Single Gift" Then
        Primary_Cell = "Contact Made"
        Secondary_Cell = "One off cash donation made immediately via CC"
        
    ElseIf Outcome_Cell = "Disconnected" Or Outcome_Cell = "Disconnected Number" Then
        Primary_Cell = "Uncontactable"
        Secondary_Cell = "Disconnected"
    ElseIf Outcome_Cell = "Deceased" Then
         Primary_Cell = "Uncontactable"
         Secondary_Cell = "Deceased"
    ElseIf Outcome_Cell = "Wrong Number" Then
        Primary_Cell = "Uncontactable"
        Secondary_Cell = "Incorrect number for donor listed"
    ElseIf Outcome_Cell = "Not Interested" Or Outcome_Cell = "Already a Supporter" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Do Not Call" Then
        Primary_Cell = "Contact Made"
        Secondary_Cell = "No RG Conversion for this campaign"
    Else
        Secondary_Cell.Interior.ColorIndex = 6
        Primary_Cell.Interior.ColorIndex = 6
    End If

End Sub

' I for individual (i.e has first name) O for organisation
Sub Generate_KeyInd(currentRow As Integer)
    Dim First_Name_Cell As Object
    Dim KeyInd_Cell As Object
    
    Set First_Name_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FirstName")))
    Set KeyInd_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("KeyInd")))
    
    If Len(Trim(First_Name_Cell)) <> 0 Then
        KeyInd_Cell = "I"
    Else
        KeyInd_Cell.Interior.ColorIndex = 6
    End If
    
End Sub
' static always 1
Sub Generate_PrimAddID(currentRow As Integer)
    Dim PrimAddID_Cell As Object
   
    Set PrimAddID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PrimAddID")))
    
    PrimAddID_Cell = "1"

End Sub
' static always 28
Sub Generate_PrimSalID(currentRow As Integer)
    Dim PrimSalID_Cell As Object
    Dim HomePh_Cell As Object
    Set PrimSalID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PrimSalID")))
    
    PrimSalID_Cell = "28"
    
End Sub

' ConsID - 1
Sub Generate_AddrImpID(currentRow As Integer)
    Dim Cons_ID As Object
    Dim AddrImpId As Object
    Dim Address_Cell As Object
    Dim Address_Country_Cell As Object
    
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set AddrImpId = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AddrImpID")))
    Set Address_Country_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AddrCountry")))
    
    
    AddrImpId = Cons_ID & "-" & "1"
    Address_Country_Cell = "Australia"
End Sub

' VICTORIA, SOUTH AUSTRALIA, NSW,  QUEENSLAND,  TASMANIA, Western Australia, ACT,  Northern Territory
Sub Generate_State(currentRow As Integer)
    Dim State_Cell As Object
    
    states = Array("", "NSW", "SA", "VIC", "WA", "QLD", "TAS", "NT", "ACT")
    PM_States = Array("", "NSW", "SOUTH AUSTRALIA", "VICTORIA", "WESTERN AUSTRALIA", "QUEENSLAND", "TASMANIA", "NORTHERN TERRITORY", "ACT")
    
    Set State_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AddrState")))
    For Each ausState In states
        If ausState = State_Cell Then
             index = Application.Match(ausState, states, False) - 1
             State_Cell = PM_States(index)
        End If
    Next ausState

End Sub


' ConsID - H - 1 only if home phone
Sub Generate_PhoneImpID1(currentRow As Integer)
    Dim Cons_ID As Object
    Dim PhoneImpID1_Cell As Object
    Dim HomePh_Cell As Object
    Dim Phone_Type_Cell As Object
     
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set PhoneImpID1_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneImpID1")))
    Set HomePh_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneNum1")))
    Set Phone_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneType1")))
    If Len(Trim(HomePh_Cell)) <> 0 Then
        PhoneImpID1_Cell = Cons_ID & "-" & "H" & "-" & "1"
        Phone_Type_Cell = "Home"
    End If
End Sub
' ConsID - M - 1 only if mobile
Sub Generate_PhoneImpID2(currentRow As Integer)
    Dim Cons_ID As Object
    Dim PhoneImpID2_Cell As Object
    Dim MobilePh_Cell As Object
    Dim Phone_Type_Cell As Object
    
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set PhoneImpID2_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneImpID2")))
    Set MobilePh_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneNum2")))
    Set Phone_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneType2")))
    If Len(Trim(MobilePh_Cell)) <> 0 Then
        PhoneImpID2_Cell = Cons_ID & "-" & "M" & "-" & "1"
        Phone_Type_Cell = "Mobile"
    End If
End Sub
'ConsID - E - 1 only if email
Sub Generate_PhoneImpID3(currentRow As Integer)
    Dim Cons_ID As Object
    Dim PhoneImpID3_Cell As Object
    Dim Email_Cell As Object
    Dim Phone_Type_Cell As Object
    
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set PhoneImpID3_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneImpID3")))
    Set Email_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneNum3")))
    Set Phone_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PhoneType3")))
    If Len(Trim(Email_Cell)) <> 0 Then
        PhoneImpID3_Cell = Cons_ID & "-" & "E" & "-" & "1"
        Phone_Type_Cell = "Email"
    End If
End Sub

'ConsID only if RG via BSB
Sub Generate_Outcome_cells(currentRow As Integer)
    Dim Cons_ID As Object
    Dim FIRRelImpID_Cell As Object
    Dim Outcome_Cell As Object
    Dim BSB_Cell As Object
    Dim BAImpID1 As Object
    Dim Bank_name As Object
    Dim GFType As Object
    Dim GFDate As Object
    Dim OutcomeUpdate As Object
    Dim GFStatusDate As Object
    Dim CampID2 As Object
    Dim GFAppeal As Object
    Dim FundID As Object
    Dim GFInsFreqNum As Object
    Dim GFInsNumDay As Object
    Dim GFInsFreqOpt As Object
    Dim GFEFT As Object
    Dim GFStatus As Object
    Dim BAImpID2 As Object
    Dim GFInsFreq As Object
    
    Dim GFSubType As Object

    
    Dim GFAttrCat1 As Object
    Dim GFAttrDate1 As Object
    Dim GFAttrDesc1 As Object
    
    Dim AppealID2 As Object
    Dim PackageID2 As Object
    Dim AppealID2Date As Object

    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set FIRRelImpID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FIRRelImpID")))
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set BSB_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BSB")))
    Set BAImpID1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BAImpID1")))
    Set Bank_name = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Bank Name")))
    Set GFType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFType")))
    Set GFDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFDate")))
    Set GFStatusDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFStatusDate")))
    Set CampID2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CampID2")))
    Set GFAppeal = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFAppeal")))
    Set FundID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("FundID")))
    Set GFInsNumDay = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFInsNumDay")))
    Set GFInsFreqNum = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFInsFreqNum")))
    Set GFInsFreqOpt = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFInsFreqOpt")))
    Set GFInsFreq = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFInsFreq")))
    
    Set GFEFT = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFEFT")))
    Set GFStatus = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFStatus")))
    Set BAImpID2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("BAImpID2")))
    
    Set GFSubType = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFSubType")))
    
    Set GFAttrCat1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFAttrCat1")))
    Set GFAttrDate1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFAttrDate1")))
    Set GFAttrDesc1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFAttrDesc1")))
    
    Set AppealID2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AppealID2")))
    Set PackageID2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("PackageID2")))
    Set AppealID2Date = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("AppealID2Date")))
    
    Set OutcomeUpdate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("OutcomeUpdateDateTime")))
    
    f = Format(OutcomeUpdate, "dd/mm/yyyy")
    
    GFSubType = "Phone - Outbound/Tele"
    AppealID2 = "RGSURCONV1617"
    PackageID2 = "WAYS"
    AppealID2Date.NumberFormat = "@"
    AppealID2Date = f
    
    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Single Gift" Then
       
         
        If Len(Trim(BSB_Cell)) <> 0 Then
            FIRRelImpID_Cell = "WT" & Cons_ID
            BAImpID1 = BSB_Cell
            Bank_name = "TRUE"
            BAImpID2 = Cons_ID
        End If
        
        If Outcome_Cell = "Confirmed" Then
            GFType = "Recurring Gift"
            GFStatusDate.NumberFormat = "@"
            GFStatusDate = f
            GFAppeal = "RGSURCONV1617"
            FundID = "58214-000"
            CampID2 = "RG-Z9031"
            GFInsFreq = "Monthly"
            GFInsNumDay = 1
            GFInsFreqNum = 5
            GFInsFreqOpt = "Specific Day"
            GFStatus = "Active"
            GFAttrCat1 = "RG Recruitment Channel"
            GFAttrDate1.NumberFormat = "@"
            GFAttrDate1 = f
            GFAttrDesc1 = "Phone"
            
        ElseIf Outcome_Cell = "Single Gift" Then
            GFType = "Cash"
            GFAppeal = "RGSURCONV1617-OO"
            GFInsNumDay = Null
            GFInsStartDate = Null
        End If
        
        GFEFT = "Yes"
        GFDate.NumberFormat = "@"
        GFDate = f
        
    End If
    
End Sub
Sub Generate_GFImpID(currentRow As Integer)
    Dim Cons_ID As Object
    Dim GFImpID_Cell As Object
    
    Set Cons_ID = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("ConsID1-PM")))
    Set GFImpID_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFImpID")))
    
    GFImpID_Cell = Cons_ID & "-" & "2"
End Sub

Sub Generate_GFInsStartDate(currentRow As Integer)
    Dim GFInsStartDate As Object
    Dim start_month As Object
    currentMonth = Month(Now)
    currentDay = Day(Now)
    
    Set GFInsStartDate = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("GFInsStartDate")))
    Set start_month = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", _
    "October", "November", "December")
    currentYear = CStr(Year(Now()))
    
    If Len(start_month) <> 0 Then
        For Each mon In longMonths
                If mon = start_month Then
                    index = Application.Match(mon, longMonths, False) - 1
                End If
        Next mon
        GFInsStartDate.NumberFormat = "@"
        GFInsStartDate = "5/" & CStr(index) & "/" & currentYear
        If Int(index) < currentMonth Then
            GFInsStartDate.Interior.ColorIndex = 6
        ElseIf Int(index) = currentMonth Then
            If currentDay >= 5 Then
                 GFInsStartDate.Interior.ColorIndex = 6
            End If
        End If
    End If
End Sub





'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF PETER MAC SPECIFIC GENERATION-------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------WAP SPECIFIC CLEANING----------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH WAP SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING             |
'|    START MONTH AND IMPORT RESULT.                                                                                   |
'|____________________________________________________________________________________________________________________|

Sub WAP_Lapsed_Generate_Debit_Date()
    Dim Outcome_Cell As Object
    Dim Debit_Date_Cell As Object
    Dim Payment_Method_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Debit_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Date")))
        Set Payment_Method_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        
    If Outcome_Cell = "Confirmed" Then
        If Payment_Method_Cell = "Credit Card" Then
            Debit_Date_Cell = "20th"
        ElseIf Payment_Method_Cell = "Direct Debit" Then
            Debit_Date_Cell = "18th"
        Else
            Debit_Date_Cell.Interior.ColorIndex = 6
        End If
    End If
    Next currentRow
End Sub

Sub Generate_Outcomes_Bank_Rejects()
    Dim Outcome_Cell As Object
    Dim RG_Cell As Object
    Dim Orig_RG_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGiftAmount")))
        Set Prev_RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnLGf_1_Amount")))
    
        newRG = CInt(RG_Cell)
        oldRG = CInt(Prev_RG_Cell)
        If Outcome_Cell = "Confirmed" Then
            If oldRG > newRG Then
                Outcome_Cell = "Confirmed/Downgraded"
                Outcome_Cell.Interior.ColorIndex = 6
                RG_Cell.INteriot.ColordIndex = 6
            End If
        End If
    Next currentRow
End Sub

Sub Generate_Report_Upgrade_Amount_ICV()
    Dim Upgrade_Cell As Object
    Dim Report_Upgrade_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Upgrade Amount")))
        Set Report_Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("UG Amount")))
        Report_Upgrade_Cell = Upgrade_Cell
        Report_Upgrade_Cell = Format(Report_Upgrade_Cell, "Currency")
    Next currentRow
End Sub


'Generates Upgrade amount by subtracting new sign up amount with previous amount.
'If upgrade amount is negative the cell is highlighted.
Sub WAP_Generate_Upgrade_Amount()
    Dim RG_Cell As Object
    Dim Prev_RG_Cell As Object
    Dim Upgrade_Cell As Object
    Dim Report_Upgrade_Cell As Object
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RegularGiftAmount")))
        Set Prev_RG_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnLGf_1_Amount")))
        Set Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Upgrade Amount")))
        Set Report_Upgrade_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("UG Amount")))
        If Trim(Len(RG_Cell)) <> 0 Then
            newAmount = CInt(RG_Cell)
            oldAmount = CInt(Prev_RG_Cell)
            upgradeAmount = newAmount - oldAmount
            Upgrade_Cell = CStr(upgradeAmount)
            Upgrade_Cell = Format(Upgrade_Cell, "Currency")
            Report_Upgrade_Cell = Upgrade_Cell
            Report_Upgrade_Cell = Format(Report_Upgrade_Cell, "Currency")
            If upgradeAmount < 0 Then
                Upgrade_Cell.Interior.ColorIndex = 6
            End If
        End If
    Next currentRow
End Sub

'Generates start month for warm wap campaigns (dd/mm/yyyy), using the existing payment method if no new
'method is found.
Sub WAP_Generate_Start_Month_Warm()
    Dim New_Month_Cell As Object
    Dim Month_Cell As Object
    Dim Payment_Cell As Object
    Dim Outcome_Cell As Object
    Dim New_Payment_Cell As Object
    Dim PaymentType As String
    
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", _
    "October", "November", "December")
    
    currentYear = CStr(Year(Now()))
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    
        Set New_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
        Set Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
        Set Payment_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("CnLGf_1_Pay_method")))
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set New_Payment_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        If Len(Trim(New_Payment_Cell)) <> 0 Then
            PaymentType = CStr(New_Payment_Cell)
        Else
            PaymentType = CStr(Payment_Cell)
        End If
        If Len(Month_Cell) <> 0 Then
            If PaymentType = "Credit Card" Then
                temp = "20/" + Month_Cell + "/" + currentYear
                New_Month_Cell = Format(temp, "dd/mm/yyyy")
            ElseIf PaymentType = "Direct Debit" Then
                temp = "18/" + Month_Cell + "/" + currentYear
                New_Month_Cell = Format(temp, "dd/mm/yyyy")
            End If
        Else
            New_Month_Cell = Null
            If Outcome_Cell = "Confirmed" Then
                New_Month_Cell.Interior.ColorIndex = 6
            End If
        End If
        Next currentRow
           
End Sub

'Generates start month for wap (dd/mm/yyyy) by using payment info and start month
Sub WAP_Generate_Start_Month_Acqusition()
    Dim New_Month_Cell As Object
    Dim Month_Cell As Object
    Dim Payment_Cell As Object
    
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", _
        "October", "November", "December")

    currentYear = CStr(Year(Now()))
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set New_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
        Set Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
        Set Payment_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Payment Method")))
        pos = InStr(Month_Cell, "/")
        If pos = 0 Then
            If Len(Month_Cell) <> 0 Then
                If Payment_Cell = "Credit Card" Then
                    temp = "20/" + Month_Cell + "/" + currentYear
                    New_Month_Cell = Format(temp, "dd/mm/yyyy")
                ElseIf Payment_Cell = "Direct Debit" Then
                    temp = "18/" + Month_Cell + "/" + currentYear
                    New_Month_Cell = Format(temp, "dd/mm/yyyy")
                End If
            Else
                New_Month_Cell = Null
            End If
        Else
            WAP_Generate_Start_Month_From_Date (currentRow)
        End If
        Next currentRow
End Sub

Sub WAP_Generate_Start_Month_From_Date(currentRow As Integer)
    Dim New_Month_Cell As Object
    Dim Month_Cell As Object
    Set Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    Set New_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Debit Day")))
    longMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", _
        "October", "November", "December")
    New_Month_Cell = Month_Cell
    New_Month_Cell.Interior.ColorIndex = 3
    monthVal = Split(Month_Cell, "/")(1)
    Month_Cell = longMonths(CInt(monthVal))
    
End Sub
'Generates import result for WAP
Sub WAP_Generate_Import_Result()
    Dim Outcome_Cell As Object
    Dim Import_Cell As Object
    
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        Set Import_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Import Result")))
        
        If Outcome_Cell = "Confirmed" Then
            Import_Cell = "Confirmed"
            
        ElseIf Outcome_Cell = "Not Interested" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Instant Refusal" _
        Or Outcome_Cell = "Already a Supporter" Or Outcome_Cell = "Do Not Call" Or Outcome_Cell = "Already a supporter" _
        Or Outcome_Cell = "Promised RG" Or Outcome_Cell = "Promised SG" Then
             Import_Cell = "Negative"
             
        ElseIf Outcome_Cell = "Uncontactable" Or Outcome_Cell = "Completed" Or Outcome_Cell = "Max Attempts" Then
             Import_Cell = "Not Available"
             
        ElseIf Outcome_Cell = "Wrong Number" Or Outcome_Cell = "Disconnected" Or Outcome_Cell = "Disconnected Number" Then
             Import_Cell = "Wrong Number"
             
        ElseIf Outcome_Cell = "Single Gift" Then
            Import_Cell = "Donation"
            
        ElseIf Outcome_Cell = "Deceased" Then
            Import_Cell = "Deceased"
            
        Else
            Import_Cell.Interior.ColorIndex = 6
        End If
    Next currentRow
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF WAP SPECIFIC GENERATION-------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------AIA SPECIFIC CLEANING----------------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO IS CONCERNED WITH AIA SPECIFIC DATA GENERATION METHODS SUCH AS GENERATING             |
'|    START MONTH AND IMPORT RESULT.                                                                                  |
'|____________________________________________________________________________________________________________________|

Sub Generate_AIA_Cash_Conversion()
    Generate_AddressLine
    Generate_Date_Of_Final_Result
    Generate_Final_Result
    Rename_Card_Type
    Generate_AIA_CC_Frequency
    Clean_And_Generate_AIA_CC_Start_Month
    Format_Expiry_Date
    Worksheets(Initiative_Name).Cells(1, "CW") = "Acq Source"
    AIA_CC_Copy_Card_Details

End Sub

Sub Generate_AIA_Lead_Conversion()
    Generate_AddressLine
    Generate_List_Type
    Generate_Date_Of_Final_Result
    Generate_Final_Result
    Generate_Source_Code
    Rename_Card_Type
    Generate_RG_Start_Date
    Format_Expiry_Date
    Format_DOB
End Sub
Sub Generate_AIA_Petition_Conversion()
    Generate_AddressLine
    Generate_Date_Of_Final_Result
    Generate_Final_Result
    'Generate_Do_Not_Call_Columns
    Rename_Card_Type
    Generate_RG_Start_Date
End Sub
Sub AIA_CC_Copy_Card_Details()
Dim Outcome_Cell As Object
Dim CC_Num_Cell As Object
Dim DD_Num_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set CC_Num_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Card Number")))
    Set DD_Num_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Account Number")))

    If Outcome_Cell = "Confirmed" Or Outcome_Cell = "Single Gift" Then
        If Len(Trim(DD_Num_Cell)) = 0 Then
            DD_Num_Cell = CC_Num_Cell
        End If
    End If
Next currentRow

End Sub

Sub Clean_And_Generate_AIA_CC_Start_Month()
Dim Outcome_Cell As Object
Dim processMonth_Cell As Object
Dim Start_Month_Cell As Object
Dim RG_Start_Date_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set processMonth_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("nextProcessmonth")))
    Set Start_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Start Month")))
    Set RG_Start_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("RG Start Date")))

    If Outcome_Cell = "Confirmed" Then
        Start = Format(RG_Start_Date_Cell, "yyyy-mm-dd")
        RG_Start_Date_Cell.NumberFormat = "@"
        RG_Start_Date_Cell = Start
        processMonth_Cell = Format(RG_Start_Date_Cell, "MMMM")
        Start_Month_Cell = processMonth_Cell
    
    Else
        RG_Start_Date_Cell = ""
    End If
Next currentRow
End Sub

Sub Generate_AIA_CC_Frequency()
Dim Outcome_Cell As Object
Dim Freq_Cell As Object
Dim Report_Freq_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
    Set Freq_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("NewFrequency")))
    Set Report_Freq_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Frequency")))
    
    If Outcome_Cell = "Confirmed" Then
        Freq_Cell = "4W"
        Report_Freq_Cell = Freq_Cell
    End If
Next currentRow

End Sub

Sub Set_Campaign(name As String)
Dim Campaign_Cell_1 As Object
Dim Campaign_Cell_2 As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Campaign_Cell_1 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Campaign")))
    Set Campaign_Cell_2 = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Campo")))
    
    Campaign_Cell_1 = name
    Campaign_Cell_2 = name
Next currentRow

End Sub

Sub Generate_Acq_Source_Recycled()
Dim Acq_Source_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Acq_Source_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Acq Source")))
    Acq_Source_Cell = "Online recycle"
Next currentRow
End Sub
Sub Format_DOB()
Dim DOB_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set DOB_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("DOB")))
    DOB_Cell = Format(DOB_Cell, "yyyy-mm-dd")
Next currentRow


End Sub

Sub Format_Expiry_Date()
Dim Expiry_Cell As Object
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Expiry_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Expiry Date")))
    If Len(Trim(Expiry_Cell)) <> 0 Then
        mon = Split(Expiry_Cell, "/")(0)
        longYear = Split(Expiry_Cell, "/")(1)
        shortYear = Mid(longYear, 3, 2)
        Expiry_Cell = mon + "/" + shortYear
    End If
Next currentRow

End Sub

Sub Generate_AddressLine()
Dim Address1_Cell As Object
Dim AddressLine_Cell As Object
addressLineCol = Column_Number("AddressLine")
address1Col = Column_Number("AddressLine1")
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set AddressLine_Cell = Worksheets(Initiative_Name).Cells(currentRow, addressLineCol)
    Set Address1_Cell = Worksheets(Initiative_Name).Cells(currentRow, address1Col)
    AddressLine_Cell = Address1_Cell
Next currentRow
End Sub

Sub Generate_List_Type()
    Dim List_Type_Cell As Object
    listTypeCol = Column_Number("List Type")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set List_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, listTypeCol)
            List_Type_Cell = "Survey"
    Next currentRow
End Sub

Sub Generate_Date_Of_Final_Result()
    Dim Date_Of_Final_Result_Cell As Object
    Dim OutcomeUpdate_Cell As Object
    outcomeUpdateCol = Column_Number("OutcomeUpdateDateTime")
    dateOfFinalResultCol = Column_Number("Date of Final Result")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Date_Of_Final_Result_Cell = Worksheets(Initiative_Name).Cells(currentRow, dateOfFinalResultCol)
        Set OutcomeUpdate_Cell = Worksheets(Initiative_Name).Cells(currentRow, outcomeUpdateCol)
            Date_Of_Final_Result_Cell.NumberFormat = "@"
            Date_Of_Final_Result_Cell = Format(OutcomeUpdate_Cell, "yyyy-mm-dd")
    Next currentRow
End Sub

Sub Generate_Final_Result()
    Dim Final_Result_Cell As Object
    Dim Outcome_Cell As Object
    finalResultCol = Column_Number("Final Result")
    outcomeCol = Column_Number("Outcome")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Final_Result_Cell = Worksheets(Initiative_Name).Cells(currentRow, finalResultCol)
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, outcomeCol)
        If Outcome_Cell = "Confirmed" Then
            Final_Result_Cell = "Confirmed"
        ElseIf Outcome_Cell = "Not Interested" Or Outcome_Cell = "Already a Supporter" Or _
        Outcome_Cell = "Instant Refusal" Or Outcome_Cell = "No Survey" Or Outcome_Cell = "Do Not Call" Then
            Final_Result_Cell = "Negative"
        ElseIf Outcome_Cell = "Uncontactable" Or Outcome_Cell = "Completed" Then
            Final_Result_Cell = "Not Available"
        ElseIf Outcome_Cell = "Wrong Number" Or Outcome_Cell = "Disconnected" Or Outcome_Cell = "Disconnected Number" Then
            Final_Result_Cell = "Wrong Number"
        ElseIf Outcome_Cell = "Deceased" Then
            Final_Result_Cell = "Deceased"
        ElseIf Outcome_Cell = "Single Gift" Then
            Final_Result_Cell = "Donation"
        Else
            Final_Result_Cell.Interior.ColorIndex = 6
        End If
    Next currentRow
End Sub


Sub Generate_Source_Code()
    Dim Source_Cell As Object
    sourceCol = Column_Number("Source")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        Set Source_Cell = Worksheets(Initiative_Name).Cells(currentRow, sourceCol)
        Source_Cell = "AQ394"
    Next currentRow
End Sub

Sub Rename_Card_Type()
 Dim Card_Type_Cell As Object
    cardTypeCol = Column_Number("Card Type")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
         Set Card_Type_Cell = Worksheets(Initiative_Name).Cells(currentRow, cardTypeCol)
         If Card_Type_Cell = "Visa" Then
            Card_Type_Cell = "VISA"
         ElseIf Card_Type_Cell = "Mastercard" Then
            Card_Type_Cell = "MC"
         ElseIf Card_Type_Cell = "Amex" Then
            Card_Type_Cell = "AMEX"
        End If
    Next currentRow
End Sub

Sub Generate_AIA_Petition_Acq_Source(Source As String)
Dim Acq_Source_Cell As Object
sourceCol = Column_Number("Acq Source")
For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
    Set Acq_Source_Cell = Worksheets(Initiative_Name).Cells(currentRow, sourceCol)
    Acq_Source_Cell = Source
Next currentRow
End Sub

Sub Generate_RG_Start_Date()
    Dim RG_Start_Date_Cell As Object
    Dim Start_Month_Cell As Object
    Dim Report_Start_Month_Cell As Object
    startMonthCol = Column_Number("nextProcessmonth")
    reportStartMonthCol = Column_Number("Start Month")
    rgStartCol = Column_Number("RG Start Date")
    longMonths = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", _
        "October", "November", "December")
    For currentRow = Sheets(1).UsedRange.Rows.Count To 2 Step -1
        isMonth = False
        Set RG_Start_Date_Cell = Worksheets(Initiative_Name).Cells(currentRow, rgStartCol)
        Set Start_Month_Cell = Worksheets(Initiative_Name).Cells(currentRow, startMonthCol)
        Set Outcome_Cell = Worksheets(Initiative_Name).Cells(currentRow, (Column_Number("Outcome")))
        If Len(Start_Month_Cell) <> 0 Then
            For Each longMonth In longMonths
                pos = InStr(longMonth, Start_Month_Cell)
                If pos <> 0 Then
                    index = Application.Match(longMonth, longMonths, False)
                    zero = ""
                    If Len(index) = 1 Then
                        zero = "0"
                    End If
                    RG_Start_Date_Cell.NumberFormat = "@"
                    RG_Start_Date_Cell = "2016-" + zero + CStr(index) + "-17"
                    RG_Start_Date_Cell.NumberFormat = "@"
                    isMonth = True
                End If
            Next longMonth
        End If
        If isMonth = False And Outcome_Cell = "Confirmed" Then
            RG_Start_Date_Cell.Interior.ColorIndex = 6
        End If
    Next currentRow
End Sub

'|--------------------------------------------------------------------------------------------------------------------|
'|-------------------------------------------END OF AIA SPECIFIC GENERATION-------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

'|====================================================================================================================|

'|--------------------------------------------------------------------------------------------------------------------|
'|-----------------------------------------GENERAL FUNCTIONS AND PROCEDURES-------------------------------------------|
'|--------------------------------------------------------------------------------------------------------------------|

' ____________________________________________________________________________________________________________________
'|                                                                                                                    |
'|    THIS SECTION OF THE MACRO CONTAINS USEFUL FUNCTIONS AND PROCEDURES THAT ARE USED THROUGHOUT THE REST            |
'|    OF THE MACRO.                                                                                                   |
'|____________________________________________________________________________________________________________________|

'takes the name of a column as a string, and returns the column number. this function is used throughout the
'entire macro as a easy way of semi-dynamically determining the correct cell.
Public Function Column_Number(name As String)
Column_Number = -1
Length = Sheets(Initiative_Name).UsedRange.Columns.Count - 1
For m = 0 To Length
    If (headerValues(m) = name) Then
        Column_Number = m + 1
    End If
Next m
If Column_Number = -1 Then
    Column_Not_Found = name
End If
End Function
'highlights the current cell yellow
Sub Highlight_Yellow()
    Current_Cell.Interior.ColorIndex = 6
End Sub
'sets the current cell to Proper Case
Sub Proper_Case()
    Current_Cell = StrConv(Current_Cell, vbProperCase)
End Sub
'sets the current cell to UPPER CASE
Sub Upper_Case()
     Current_Cell = StrConv(Current_Cell, vbUpperCase)
End Sub
'sets the current cell as text
Sub Set_As_Text()
    Current_Cell.NumberFormat = "@"
End Sub
'sets the current cell as currency
Sub Set_As_Currency()
    Current_Cell = Format(Current_Cell, "Currency")
End Sub
'this is the error handling method for the single column cleaning section. it will tell you which
'column the macro encountered an error on, and will highlight the last cell cleaned red.
Sub Handle_Error(errLocation As String)
    Current_Cell.Interior.ColorIndex = 3
    Error_Has_Occurred = True
    strPrompt = "There has been an error in column: " + colName
    strTitle = "Error in " + errLocation
    iRet = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
End Sub
'this is the error handling method used when initializing the payment details to be used to aid the cleaning / generation
'process. This will tell you which column mapping is incorrect, in order to easily change in the campaign sheets.
Sub Handle_Initialize_Failure()
    Error_Has_Occurred = True
    strPrompt = "Error Initializing " + Column_Not_Found + "! Please check the column mapping."
    iRet = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
End Sub
'Handles any erros in the core data generation logic that is common to most charities.
Sub Handle_Core_Data_Failure()
    Error_Has_Occurred = True
    strPrompt = "Error in " + Column_Not_Found + " Data Generation Procedure! Please check Data Generation Logic."
    iRet = MsgBox(strPrompt, vbOKOnly + vbInformation, strTitle)
End Sub
'highlights the top row if there is any highlighted cells in the associated column.
Sub Highlight_Headers()
    rowCount = Worksheets(Initiative_Name).UsedRange.Rows.Count
    colCount = Worksheets(Initiative_Name).UsedRange.Columns.Count
    For i = colCount To 1 Step -1
        For j = rowCount To 2 Step -1
            If Worksheets(Initiative_Name).Cells(j, i).Interior.ColorIndex = 6 Then
               Worksheets(Initiative_Name).Cells(1, i).Interior.ColorIndex = 6
            End If
        Next j
    Next i
End Sub



















