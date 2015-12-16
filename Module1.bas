Option Compare Database
Option Explicit
' Added 12/16/2015
Function ComboKeyValueStore(CtrlName As Control)
Dim strActiveComboValue As String

    strActiveComboValue = CtrlName.Column(1)
    TempVars.Add "ActiveComboValue", strActiveComboValue

'MsgBox TempVars!ActiveComboValue


End Function

Function LogChangesOriginal(lngID As Long, Optional strField As String = "")
'    Dim dbs As DAO.Database
'    Dim rst As DAO.Recordset
'    Dim varOld As Variant
'    Dim varNew As Variant
'    Dim strFormName As String
'    Dim strControlName As String
'
'    varOld = Screen.ActiveControl.OldValue
'    varNew = Screen.ActiveControl.Value
'    strFormName = Screen.ActiveForm.Name
'    strControlName = Screen.ActiveControl.Name
'    Set dbs = CurrentDb()
'    Set rst = dbs.TableDefs("ztblDataChanges").OpenRecordset
'
'    With rst
'        .AddNew
'        !FormName = strFormName
'        !ControlName = strControlName
''        If strField = "" Then
''            !FieldName = strControlName
''        Else
''            !FieldName = strField
''        End If
'        !RecordID = lngID
'        !UserName = Environ("username")
'        If Not IsNull(varOld) Then
'            !OldValue = CStr(varOld)
'        End If
'        !newValue = CStr(varNew)
'        .Update
'    End With
'    'clean up
'    rst.Close
'    Set rst = Nothing
'    dbs.Close
'    Set dbs = Nothing
'
'
End Function


Function LogChanges(lngID As Long)
'    Dim dbs As DAO.Database
'    Dim rst As DAO.Recordset
'    Dim varOld As Variant
'    Dim varNew As Variant
'    Dim strFormName As String
'    Dim strControlName As String
'    Dim strControlSource As String
'    Dim strControlParentName As String
'    Dim strControlParentRecordSource As String
'    Dim lngControlType As Long
'    Dim lngVarType As Long
'
'
'    varNew = Screen.ActiveControl.Value
'    Debug.Print varNew
'
''    If Screen.ActiveControl.OldValue = Null Then ' Assume no old value means record did not exist
''        varOld = varNew
''    Else
'    varOld = Screen.ActiveControl.OldValue
''    End If
'
'    strFormName = Screen.activeForm.Name
'    strControlName = Screen.ActiveControl.Name
'    strControlSource = Screen.ActiveControl.ControlSource
'    strControlParentName = Screen.ActiveControl.Parent.Name
'    'strControlParentRecordSource = Screen.ActiveControl.Parent.RecordSource
'    lngControlType = Screen.ActiveControl.ControlType
'    lngVarType = VarType(varNew)
'
'If varOld = varNew Then
' 'No changes made, do not write data
'Else
'    Set dbs = CurrentDb()
'    Set rst = dbs.TableDefs("ztblDataChanges").OpenRecordset
'
'    With rst
'        .AddNew
'        If strFormName = "Main" Then
'        !FormName = Forms!Main.NavigationSubform.SourceObject
'        Else
'        !FormName = strFormName
'        End If
'
'
'        !ControlName = strControlName
'        '!ControlParentRecordSource = strControlParentRecordSource
'        !ControlParentName = strControlParentName
'        !ControlSource = strControlSource
'        !ControlType = lngControlType
'        !DataType = lngVarType
'        !RecordID = lngID
'        !UserName = Environ("username")
'
'        Select Case lngControlType
'            Case Is = 111 'ComboBox
'                !NewBoundValue = Screen.ActiveControl.Column(Screen.ActiveControl.BoundColumn)
'                !OldBoundValue = TempVars!ActiveComboValue
'                !newValue = CStr(varNew)
'                !OldValue = CStr(varOld)
'
'            Case Else
'                !NewBoundValue = CStr(varNew)
'                !OldBoundValue = CStr(varOld)
'                !newValue = ""
'                !OldValue = ""
'        End Select
'
'        .Update
'    End With
'
'    'clean up
'    rst.Close
'    Set rst = Nothing
'    dbs.Close
'    Set dbs = Nothing
'End If
'
'Screen.activeForm.Requery
'
'End Function


''''AcControlType for Reference:
''''
''''Constant Value
''''acBoundObjectFrame 108
''''acCheckBox 106
''''acComboBox 111
''''acCommandButton 104
''''acCustomControl 119
''''acImage 103
''''acLabel 100
''''acLine 102
''''acListBox 110
''''acObjectFrame 114
''''acOptionButton 105
''''acOptionGroup 107
''''acPage 124
''''acPageBreak 118
''''acRectangle 101
''''acSubform 112
''''acTabCtl 123
''''acTextBox 109
''''acToggleButton 122


'''

''''VarType for Reference:
''''
''''The required varname argument (argument: A value that provides information to an action, an event, a method, a property, a function, or a procedure.) is a Variant (Variant data type: The default data type for variables that don't have type-declaration characters when a Deftype statement isn't in effect. A Variant can store numeric, string, date/time, Null, or Empty data.) containing any variable except a variable of a user defined type (user-defined type: In VBA, any data type defined using the Type statement. User-defined data types can contain one or more elements of any data type. Arrays of user-defined and other data types are created using the Dim statement.).
''''
''''Return Values
''''
''''Constant Value Description
''''vbEmpty 0 Empty (Empty: The state of an uninitialized Variant variable (which returns a VarType of 0). Not to be confused with Null (a variable state indicating invalid data), variables with zero-length strings (" "), or numeric variables equal zero.) (uninitialized)
''''vbNull 1 Null (Null: A value you can enter in a field or use in expressions or queries to indicate missing or unknown data. In Visual Basic, the Null keyword indicates a Null value. Some fields, such as primary key fields, can't contain a Null value.) (no valid data)
''''vbInteger 2 Integer
''''vbLong 3 Long integer
''''vbSingle 4 Single-precision floating-point number
''''vbDouble 5 Double-precision floating-point number
''''vbCurrency 6 Currency value
''''vbDate 7 Date value
''''vbString 8 String
''''vbObject 9 Object
''''vbError 10 Error value
''''vbBoolean 11 Boolean value
''''vbVariant 12 Variant (used only with arrays (array: A variable that contains a finite number of elements that have a common name and data type. Each element of an array is identified by a unique index number. Changes made to one element of an array don't affect the other elements.) of variants)
''''vbDataObject 13 A data access object
''''vbDecimal 14 Decimal value
''''vbByte 17 Byte value
''''vbUserDefinedType 36 Variants that contain user-defined types
''''vbArray 8192 Array