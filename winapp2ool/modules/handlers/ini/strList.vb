Public Class strList
    Private lst As List(Of String)

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
    Public Function contains(givenValue As String, Optional ignoreCase As Boolean = False)
        If ignoreCase Then
            For Each value In items
                If givenValue.Equals(value, StringComparison.InvariantCultureIgnoreCase) Then Return True
            Next
            Return False
        Else
            Return items.Contains(givenValue)
        End If
    End Function
End Class
