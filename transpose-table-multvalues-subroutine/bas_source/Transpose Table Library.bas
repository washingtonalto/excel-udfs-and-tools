Attribute VB_Name = "TransposeLibrary"
' Author            : Washington Alto
' Copyright       : Copyright (c) 2018 Washington Alto. All rights reserved
' Date Modified : October 1, 2018
' Version           : 1
' List                  :
'       Public Sub localTransposeMultValueFieldTable(srcWorksheetName As String, srcTableName As String, dstWorksheetName As String, _
'                   dstTableName As String, transposeField As String, Optional dstTablePlacementRange As String = "A1", Optional strDelim As String = ";")
'       Public Function localGetStringofInclusiveDates(dteStart As Date, dteEnd As Date, Optional strDelim As String = ";", Optional maxDayslimit As Integer = 365) As String

Public Sub localTransposeMultValueFieldTable(srcWorksheetName As String, srcTableName As String, dstWorksheetName As String, _
dstTableName As String, transposeField As String, Optional dstTablePlacementRange As String = "A1", Optional strDelim As String = ";")
' Sub localTransposeMultValueFieldTable - This routine creates a destination Excel table based on the source Excel table where delimited values on transposeField
'     is transposed for every row and other columns are the same for each row
' Parameters:
'           srcWorksheetName - Source Worksheet Name
'           srcTableName - Source Table Name located in Source Worksheet
'           dstWorksheetName - Destination Worksheet Name
'           dstTableName - Destination Table Name located in Destination Worksheet
'           transposeField - Source Field Name to be delimited via delimiter and transpose
'           dstTablePlacementRange - Range where to position the Destination Table on the Destination Worksheet
'           strDelim - delimiter string used for delimiting strings on the transposeField
'
'  Example:
'           Call localTransposeMultValueFieldTable("Source 1", "Source_Table", "Source 1", "Destination_Table", "Month", "H3")

Dim srcWorksheet, dstWorksheet As Worksheet
Dim srcTable, dstTable As ListObject
Dim srcListRow, dstListRow As ListRow
Dim srcListRowCount, srcListColumnCount, rowCtr, colCtr As Integer

Set srcWorksheet = Sheets(srcWorksheetName)
Set dstWorksheet = Sheets(dstWorksheetName)

Set srcTable = srcWorksheet.ListObjects(srcTableName)
' Delete destination table if it exist
For Each dstTable In dstWorksheet.ListObjects
    If dstTable.Name = dstTableName Then
            dstTable.Delete
    End If
Next dstTable
' Create destination table
Set dstTable = dstWorksheet.ListObjects.Add(xlSrcRange, dstWorksheet.Range(dstTablePlacementRange), , xlYes)
dstTable.Name = dstTableName

' Copy the Table Header from Src to Dst
srcTable.HeaderRowRange.Copy Destination:=dstWorksheet.Range(dstTableName + "[[#Headers],[Column1]]")

' Get the row and column count of the Source Table
srcListRowCount = srcTable.ListRows.Count
srcListColCount = srcTable.ListColumns.Count

' Get the column position number for the transposeField
Dim postransposeField As Integer
For colCtr = 1 To srcListColCount
        If srcTable.ListColumns(colCtr).Name = transposeField Then
            postransposeField = colCtr
        End If
Next colCtr

' Define variables that will be used in the routine for counting the number of delimiters within the transposeField
Dim strtransposeFieldLen As Integer
Dim numtransposeFieldCtr As Integer
Dim chartransposeFieldCtr As Integer
Dim strtransposeField As String

For rowCtr = 1 To srcListRowCount

        ' strtransposeField will contain the value of the transposeField for that particular row
        strtransposeField = srcTable.DataBodyRange(rowCtr, postransposeField).Value
        strtransposeFieldLen = Len(strtransposeField)
        chartransposeFieldCtr = 0
        For numtransposeFieldCtr = 1 To strtransposeFieldLen
            If Mid(strtransposeField, numtransposeFieldCtr, Len(strDelim)) = strDelim Then
                chartransposeFieldCtr = chartransposeFieldCtr + 1
            End If
        Next numtransposeFieldCtr
        '   chartransposeFieldCtr should now contain the number of delimiters within the transposeField for that particular row
        
        ' Defined an array of string for the transposeFieldValue for the particular row
        Dim arrtransposeFieldVal() As String
        arrtransposeFieldVal = Split(strtransposeField, strDelim, chartransposeFieldCtr + 1, vbTextCompare)

        Dim transposeFieldIndividualVal As Variant
        For Each transposeFieldIndividualVal In arrtransposeFieldVal
                Set dstListRow = dstTable.ListRows.Add(AlwaysInsert:=True)
                For colCtr = 1 To srcListColCount
                        If colCtr = postransposeField Then
                                dstListRow.Range.Columns(colCtr).Value = Trim(transposeFieldIndividualVal)
                        Else
                                dstListRow.Range.Columns(colCtr).Value = srcTable.DataBodyRange(rowCtr, colCtr).Value
                        End If
                Next colCtr
        Next transposeFieldIndividualVal
        
        ' Erase the array of transposeFieldVal to release memory used by the array before the next iteration
        Erase arrtransposeFieldVal
Next rowCtr

' Release object memory
Set srcWorksheet = Nothing
Set dstWorksheet = Nothing
Set srcTable = Nothing
Set dstTable = Nothing
End Sub

Public Function localGetStringofInclusiveDates(dteStart As Date, dteEnd As Date, Optional strDelim As String = ";", Optional maxDayslimit As Integer = 365) As String
' Sub localGetStringofInclusiveDates - This user-defined function returns a string of dates separated by strDelim from Start Date dteStart to End Date dteEnd
' Parameters:
'       dteStart - Start Date
'       dteEnd - End Date
'       strDelim - Delimiter string
'       maxDayslimit - if specified, the maximum number of days between start and end date to prevent unnecessarily long date string returned
' Example:
'       localGetStringofInclusiveDates([@[Start Date]],[@[End Date]])
'       localGetStringofInclusiveDates([@[Start Date]],[@[End Date]],",")
'       localGetStringofInclusiveDates([@[Start Date]],[@[End Date]],",",40)
' Sample output: 2018-01-03;2018-01-04;2018-01-05 where start date is Jan 3, 2018 and end date is Jan 5, 2018
'

Dim dteStartDateOnly, dteEndDateOnly As Date
Dim strOutputString As String
Dim numDaysBetweenInterval As Integer
Dim numDayCtr As Integer
Dim dteDateStr As String

strOutputString = ""
' The statements below forces the function to ignore the time part of the date if they exist
dteStartDateOnly = DateValue(Format(dteStart, "yyyy/mm/dd"))
dteEndDateOnly = DateValue(Format(dteEnd, "yyyy/mm/dd"))

If dteStartDateOnly > dteEndDateOnly Then
        strOutputString = ""
ElseIf dteStartDateOnly = dteEndDateOnly Then
        strOuputString = Format(dteStartDateOnly, "yyyy-mm-dd")
Else
        numDaysBetweenInterval = DateDiff("d", dteStartDateOnly, dteEndDateOnly, vbUseSystemDayOfWeek)
        If numDaysBetweenInterval > maxDayslimit Then
                numDaysBetweenInterval = maxDayslimit - 1
        End If
        For numDayCtr = 0 To numDaysBetweenInterval
                dteDateStr = Format(DateAdd("d", numDayCtr, dteStartDateOnly), "yyyy-mm-dd", vbUseSystemDayOfWeek)
                strOuputString = strOuputString + IIf(numDayCtr = numDaysBetweenInterval, dteDateStr, dteDateStr + strDelim)
        Next numDayCtr
End If


localGetStringofInclusiveDates = strOuputString

End Function

