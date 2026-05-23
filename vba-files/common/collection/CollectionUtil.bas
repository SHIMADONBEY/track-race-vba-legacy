Attribute VB_Name = "CollectionUtil"
'namespace=vba-files/common/collection
Option Explicit
Option Private Module

Private Const MIN_SIZE_TO_MERGE As Long = 32

Public Sub SwapItem(ByRef List As Collection, ByVal LeftIndex As Long, ByVal RightIndex As Long)
    Dim leftHand As Object: Set leftHand = List.Item(LeftIndex)
    Dim rightHand As Object: Set rightHand = List.Item(RightIndex)

    Call List.Add(rightHand, After:=LeftIndex)
    Call List.Remove(LeftIndex)
    Call List.Add(leftHand, After:=RightIndex)
    Call List.Remove(RightIndex)
End Sub

'/*
'# CollectionUtil.SortObjectCollection()
'
'Sorts the specified list according to the order induced by the specified comparator.
'This sort is guaranteed to be stable: equal elements will not be reordered as a result of the sort.
'
'** Syntax **
' List           - The list to sort.
' Comparator     - The comparator to determine the order of the list.
' Descending     - If the value is set to True, sort in descending order.
'
'** Returns **
' Sorted List.
'*/
Public Function SortObjectCollection(ByRef List As Collection, Comparator As IObjectComparator, Optional Descending As Boolean = False) As Collection
    Dim arraySize As Long: arraySize = List.Count
    Dim minToRun As Long: minToRun = MinRunLength(MIN_SIZE_TO_MERGE)
    Dim itemArray() As Object: itemArray = ConvertToArray(List)
    Dim indexes(3) As Long

    indexes(0) = 0
    Dim i As Long
    Dim l As Long: l = 0
    For i = 1 To 3
        l = l + CreateRun(itemArray, Comparator, Descending, l, Application.WorksheetFunction.Min(MIN_SIZE_TO_MERGE - 1, arraySize - l))
        indexes(i) = l
    Next i

    Do While indexes(1) + 1 < arraySize
        Dim length_0 As Long
        Dim length_1 As Long
        Dim length_2 As Long

        Do While True
            length_0 = indexes(1) - indexes(0)
            length_1 = indexes(2) - indexes(1)
            length_2 = indexes(3) - indexes(2)

            If (length_0 > length_1 + length_2) Then
                If (length_1 > length_2) Then
                    Exit Do
                Else
                    Call MergeItems(itemArray, indexes(0), indexes(1), indexes(2) - 1, Comparator, Descending)
                    indexes(1) = indexes(2)
                    indexes(2) = indexes(3)
                End If
            Else
                If (length_0 <= length_2) Then
                    Call MergeItems(itemArray, indexes(0), indexes(1), indexes(2) - 1, Comparator, Descending)
                    indexes(1) = indexes(2)
                    indexes(2) = indexes(3)
                Else
                    If length_2 < 1 Then Exit Do
                    Call MergeItems(itemArray, indexes(1), indexes(2), indexes(3) - 1, Comparator, Descending)
                    indexes(2) = indexes(3)
                End If
            End If

            indexes(3) = indexes(3) + CreateRun(itemArray, Comparator, Descending, indexes(3), Application.WorksheetFunction.Min(MIN_SIZE_TO_MERGE - 1, arraySize - indexes(3)))
        Loop

        If (length_2 > 0) Then
            Call MergeItems(itemArray, indexes(1), indexes(2), indexes(3) - 1, Comparator, Descending)
            indexes(2) = indexes(3)
        End If

        If (length_1 > 0) Then
            Call MergeItems(itemArray, indexes(0), indexes(1), indexes(2) - 1, Comparator, Descending)
            indexes(1) = indexes(2)
        End If

        l = indexes(1)
        For i = 2 To 3
            l = l + CreateRun(itemArray, Comparator, Descending, l, Application.WorksheetFunction.Min(MIN_SIZE_TO_MERGE - 1, arraySize - l))
            indexes(i) = l
        Next i
    Loop

    Set SortObjectCollection = ConvertToCollection(itemArray)
End Function

'/*
'# CollectionUtil.GroupingBy()
'
'This key maps to a collection keyed by the values resulting from applying the classification function to the input elements.
'
'** Syntax **
' List          - The list to group.
' Classifier  - The classifier function mapping input elements to keys.
'
'** Returns **
' Implements a "group by" operation on an object list, grouping the elements according to the classification function and returning the result in a Dictionary.
'*/
Public Function GroupingBy(ByRef List As Collection, Classifier As IObjectClassifier) As Object
    Dim recordItem As Object
    Dim groupMap As Object: Set groupMap = CreateObject("Scripting.Dictionary")

    For Each recordItem In List
        Dim itemKey As Variant: itemKey = Classifier.Classify(recordItem)
        Dim itemsList As Collection
        If Not (groupMap.Exists(itemKey)) Then
            Call groupMap.Add(itemKey, New Collection)
        End If

        Set itemsList = groupMap.Item(itemKey)
        Call itemsList.Add(recordItem)
        Set groupMap.Item(itemKey) = itemsList
    Next recordItem 

    Set GroupingBy = groupMap
End Function

Public Function ConvertToArray(Items As Collection) As Variant
    Dim itemArray() As Object
    Dim element As Object
    Dim itemCount As Long: itemCount = Items.Count
    Dim index As Long

    If itemCount > 0 Then
        ReDim itemArray(0 To itemCount - 1)
    End If

    For index = 1 To itemCount
        Set itemArray(index - 1) = Items.item(index)
    Next index

    ConvertToArray = itemArray
End Function

Public Function ConvertToCollection(Items As Variant) As Collection
    Dim itemCollection As Collection: Set itemCollection = New Collection
    Dim index As Long

    For index = LBound(Items) To UBound(Items)
        Call itemCollection.Add(Items(index))
    Next index

    Set ConvertToCollection = itemCollection
End Function

Private Function MinRunLength(size As Long) As Long
    Debug.Assert size >= 0

    Dim r As Long: r = 0
    Dim n As Long: n = size
    Do While n >= MIN_SIZE_TO_MERGE
        r = r Or (size And 1)
        n = n \ (2 ^ 1)
    Loop

    MinRunLength = n + r
End Function

Private Function CreateRun(Items() As Object, Comparator As IObjectComparator, Descending As Boolean, LeftIndex As Long, Optional MaxSize As Long = (MIN_SIZE_TO_MERGE - 1)) As Long
    If MaxSize < 2 Then
        CreateRun = MaxSize
        Exit Function
    End If

    Dim i As Long
    Dim minSizeRun As Long: minSizeRun = MinRunLength(MIN_SIZE_TO_MERGE)
    Dim stoppedAt As Long

    If (minSizeRun > MaxSize) Then 
        stoppedAt = LeftIndex + MaxSize - 1
    Else
        For i = minSizeRun To MaxSize - 1
            If (Comparator.Compare(Items(LeftIndex + i - 1), Items(LeftIndex + i)) * IIf(Descending, -1, 1) > 0) Then
                Exit For
            End If
        Next i
        stoppedAt = LeftIndex + i - 1
    End If

    Call BinaryInsertionSort(Items, LeftIndex, stoppedAt, Comparator, Descending)
    CreateRun = stoppedAt - LeftIndex + 1
End Function

Private Sub BinaryInsertionSort(Items() As Object, LeftIndex As Long, RightIndex As Long, Comparator As IObjectComparator, Descending As Boolean)
    Dim i As Long, j As Long
    Dim temp As Object

    For i = LeftIndex + 1 To RightIndex
        Set temp = Items(i)
        If (Comparator.Compare(temp, Items(i - 1)) * IIf(Descending, -1, 1) < 0) Then
            Dim index1 As Long: index1 = LeftIndex
            Dim index2 As Long: index2 = i - 1
            Dim cursor As Long: cursor = index1
            
            Do While ((index2 - index1) > 0)
                cursor = Fix((index1 + index2) / 2)
                If (Comparator.Compare(temp, Items(cursor)) * IIf(Descending, -1, 1) < 0) Then
                    index2 = cursor
                Else
                    If index2 - index1 = 1 Then
                        cursor = cursor + 1
                    End If
                    index1 = cursor
                End If
            Loop

            For j = i To cursor + 1 Step -1
                Set Items(j) = Items(j - 1)
            Next j
            Set Items(cursor) = temp
        End If
    Next i
End Sub

Private Sub MergeItems(Items() As Object, LeftIndex As Long, MiddleIndex As Long, RightIndex As Long, Comparator As IObjectComparator, Descending As Boolean)
    Dim leftLength As Long: leftLength = MiddleIndex - LeftIndex
    Dim rightLength As Long: rightLength = RightIndex - MiddleIndex + 1
    Dim widerLength As Long
    Dim shorterLength As Long
    Dim tempArray() As Object
    Dim arrayShifted As Boolean
    Dim x As Long

    arrayShifted = (leftLength > rightLength)
    If (arrayShifted) Then 
        widerLength = leftLength
        shorterLength = rightLength
        ReDim tempArray(0 To shorterLength - 1)

        For x = 0 To shorterLength - 1
            Set tempArray(x) = Items(MiddleIndex + x)
        Next x

        ' Shift the array to the right.
        For x = widerLength - 1 To 0 Step -1
            Set Items(LeftIndex + x + shorterLength) = Items(LeftIndex + x)
        Next x
    Else
        widerLength = rightLength
        shorterLength = leftLength
        ReDim tempArray(0 To shorterLength - 1)

        For x = 0 To shorterLength - 1
            Set tempArray(x) = Items(LeftIndex + x)
        Next x
    End If

    Dim i As Long: i = RightIndex - widerLength + 1
    Dim j As Long: j = 0
    Dim k As Long: k = LeftIndex

    Do While (i <= RightIndex And j < shorterLength)
        Dim compared As Integer: compared = Comparator.Compare(Items(i), tempArray(j)) * IIf(Descending, -1, 1)
        If (compared = 0) Then 
            If (arrayShifted) Then 
                Set Items(k) = Items(i)
                i = i + 1
            Else
                Set Items(k) = tempArray(j)
                j = j + 1
            End If
        ElseIf (compared < 0) Then
            Set Items(k) = Items(i)
            i = i + 1
        Else
            Set Items(k) = tempArray(j)
            j = j + 1
        End If

        k = k + 1
    Loop

    Do While (j < shorterLength)
        Set Items(k) = tempArray(j)
        k = k + 1
        j = j + 1
    Loop
End Sub
