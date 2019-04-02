Public Class strList
    Private lst As New List(Of String)

    ''' <summary>
    ''' A list of String values
    ''' </summary>
    ''' <returns></returns>
    Public Property items As List(Of String)
        Get
            Return lst
        End Get
        Set(value As List(Of String))
            lst = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the index of the given String in the list if it exists. Else -1
    ''' </summary>
    ''' <param name="item">A String to search for in the list</param>
    ''' <returns></returns>
    Public Function indexOf(item As String) As Integer
        Return lst.IndexOf(item)
    End Function

    ''' <summary>
    ''' Conditionally adds an item to the list 
    ''' </summary>
    ''' <param name="item">A string value to add to the list</param>
    ''' <param name="cond">The optional condition under which the value should be added (default: true)</param>
    Public Sub add(item As String, Optional cond As Boolean = True)
        If cond Then lst.Add(item)
    End Sub

    ''' <summary>
    ''' Returns true if the list contains a given value. Case sensitive by default
    ''' </summary>
    ''' <param name="givenValue">A value to search the list for</param>
    ''' <param name="ignoreCase">The optional condition specifying whether string casing should be ignored</param>
    ''' <returns></returns>
    Public Function contains(givenValue As String, Optional ignoreCase As Boolean = False) As Boolean
        If ignoreCase Then
            For Each value In items
                If givenValue.Equals(value, StringComparison.InvariantCultureIgnoreCase) Then Return True
            Next
            Return False
        Else
            Return items.Contains(givenValue)
        End If
    End Function

    ''' <summary>
    ''' Construct a list of neighbors for strings in a list
    ''' </summary>
    Public Function getNeighborList() As List(Of KeyValuePair(Of String, String))
        Dim neighborList As New List(Of KeyValuePair(Of String, String))
        If lst.Count > 1 Then

        End If
        neighborList.Add(New KeyValuePair(Of String, String)("first", lst(1)))
        For i = 1 To lst.Count - 2
            neighborList.Add(New KeyValuePair(Of String, String)(lst(i - 1), lst(i + 1)))
        Next
        neighborList.Add(New KeyValuePair(Of String, String)(lst(lst.Count - 2), "last"))
        Return neighborList
    End Function

    ''' <summary>
    ''' Replaces an item in a list of strings at the index of another given item
    ''' </summary>
    ''' <param name="indexOfText">The text to be replaced</param>
    ''' <param name="newText">The replacement text</param>
    ''' strlist candidate
    Public Sub replaceStrAtIndexOf(indexOfText As String, newText As String)
        items(items.IndexOf(indexOfText)) = newText
    End Sub

    ''' <summary>
    ''' Returns the number of items in the strList
    ''' </summary>
    ''' <returns></returns>
    Public Function count() As Integer
        Return items.Count
    End Function

End Class
