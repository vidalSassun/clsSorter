VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'variables
Private oTerms As Object, oMasks As Object, oExceptionMasks As Object, oHistory As Object
Public delimiter As String, exceptionWord As String

'object initilization
Private Sub Class_Initialize()
    'default delimiter is ","
    delimiter = ","
    'default exception word is "EXCEPT"
    exceptionWord = "EXCEPT"
    'creating "Dictionary" objects
    Set oTerms = CreateObject("Scripting.Dictionary")
    Set oMasks = CreateObject("Scripting.Dictionary")
    Set oExceptionMasks = CreateObject("Scripting.Dictionary")
    Set oHistory = CreateObject("Scripting.Dictionary")
End Sub

'object termination
Private Sub Class_Terminate()
    Set oTerms = Nothing
    Set oMasks = Nothing
    Set oExceptionMasks = Nothing
    Set oHistory = Nothing
End Sub

'push function
'adds a new element in an array
'paramaters:
'vItem - element to add
'aArray - array, if missed creates a new one
'return:
'an array with added element
Private Function pushItem(ByVal vItem, Optional ByVal aArray As Variant = Empty) As Variant
    Dim lLast As Long
    If IsArray(aArray) = False Then
        'if aArray argument is missed, creates a new array containing the only element - vItem
        If aArray = Empty Then
            pushItem = Array(vItem)
            Exit Function
        'else uses aArray argument
        Else
            aArray = Array(aArray)
        End If
    End If
    'adds an element vItem to array
    lLast = UBound(aArray) + 1
    ReDim Preserve aArray(LBound(aArray) To lLast)
    aArray(lLast) = vItem
    pushItem = aArray
End Function

'method adds marking conditions
'parameters:
'sMarker - marker
'arrTerms - sorting conditions
Public Sub addTerms(ByRef sMarker As String, ByRef arrTerms As Variant)
    'variables
    Dim vItem As Variant, vMask As Variant, vExceptionMask As Variant, vTemp As Variant, vSplittedTerms As Variant, vSplittedItem As Variant
    Dim iIndex As Integer
    'check if marker conditions exists, rewrite if exists
    If oHistory.Exists(sMarker) Then deleteTerms sMarker
    'conditions indexed starting from 0
    iIndex = 0
    'creates vMask and vExceptionMask elements
    For Each vItem In arrTerms
        'if vItem is empty, adds False to masks
        If vItem = Empty Or vItem = "" Then
            vMask = pushItem(False, vMask)
            vExceptionMask = pushItem(False, vExceptionMask)
            vSplittedTerms = pushItem(Empty, vSplittedTerms)
        'else adds True to vMask and checkes if condition body contains exception word
        Else
            vMask = pushItem(True, vMask)
            'if condition contains exception word
            If vItem Like "*" & exceptionWord & "*" Then
                'cuts exception word from condition body
                vTemp = InStr(1, vItem, exceptionWord)
                vItem = Left(vItem, vTemp - 1) & Mid(vItem, vTemp + Len(exceptionWord))
                'adds True to vExceptionMask
                vExceptionMask = pushItem(True, vExceptionMask)
            Else
                'adds False to vExceptionMask
                vExceptionMask = pushItem(False, vExceptionMask)
            End If
            'splits condition on several items
            vItem = Split(vItem, delimiter)
            vSplittedItem = Empty
            For Each vTemp In vItem
                'deletes unnecessary spaces
                vTemp = Trim(vTemp)
                vSplittedItem = pushItem(vTemp, vSplittedItem)
                With oTerms
                    'if condition with current index doesn't exist, creates a new dictionary object
                    If .Exists(iIndex) = False Then Set .Item(iIndex) = CreateObject("Scripting.Dictionary")
                    With .Item(iIndex)
                        'if item doesn't exist, creates a new dictionary object
                        If .Exists(vTemp) = False Then Set .Item(vTemp) = CreateObject("Scripting.Dictionary")
                        With .Item(vTemp)
                        .Item(sMarker) = Empty
                        End With
                    End With
                End With
            Next
            vSplittedTerms = pushItem(vSplittedItem, vSplittedTerms)
        End If
        'next condition index
        iIndex = iIndex + 1
    Next
    'adds elements to oMasks, oExceptionMasks, oHistory
    oMasks.Item(sMarker) = vMask
    oExceptionMasks.Item(sMarker) = vExceptionMask
    oHistory.Item(sMarker) = vSplittedTerms
End Sub

'method returns marker based on conditions
'parameter:
'arrTerms - conditions array
'return:
'marker, if exists, else - empty value
Public Function getMarker(ByRef arrTerms As Variant) As String
    'variables
    Dim vItem As Variant, vMask As Variant, vExceptionMask As Variant, vMarker As Variant, vTemp As Variant
    Dim iIndex As Integer
    Dim bFound As Boolean
    'finds marker using masks
    For Each vMarker In oMasks.Keys()
        vMask = oMasks.Item(vMarker)
        vExceptionMask = oExceptionMasks.Item(vMarker)
        With oTerms
            For Each vItem In .Keys()
                'if "excluding" condition
                If vMask(vItem) And vExceptionMask(vItem) Then
                    With .Item(vItem)
                        vItem = arrTerms(vItem)
                        If .Exists(vItem) Then
                            With .Item(vItem)
                                'checkes if condition for this marker exists
                                If .Exists(vMarker) Then
                                    bFound = False
                                    Exit For
                                Else
                                    bFound = True
                                End If
                            End With
                        Else
                            bFound = True
                        End If
                    End With
                'if "including" condition
                ElseIf vMask(vItem) Then
                    With .Item(vItem)
                        'checkes if item exists in "including" conditions
                        vItem = arrTerms(vItem)
                        If .Exists(vItem) Then
                            With .Item(vItem)
                                'checkes if condition for this marker doesn't exist
                                If .Exists(vMarker) Then
                                    bFound = True
                                Else
                                    bFound = False
                                    Exit For
                                End If
                            End With
                        Else
                            bFound = False
                            Exit For
                        End If
                    End With
                End If
            Next
        End With
        If bFound Then
            getMarker = vMarker
            Exit Function
        End If
    Next
End Function

'method deletes existing terms
'parameters:
'sMarker - marker
Public Sub deleteTerms(ByRef sMarker As String)
    'variables
    Dim vTerm As Variant, vItem As Variant
    Dim iIndex As Integer
    'looks through oTerms items and deletes every mentioning of sMarker
    With oHistory
        If .Exists(sMarker) Then
            iIndex = 0
            For Each vTerm In .Item(sMarker)
                For Each vItem In vTerm
                    With oTerms.Item(iIndex)
                        With .Item(vItem)
                            .Remove sMarker
                        End With
                    End With
                Next
                iIndex = iIndex + 1
            Next
        End If
        .Remove sMarker
    End With
    oMasks.Remove sMarker
    oExceptionMasks.Remove sMarker
End Sub
