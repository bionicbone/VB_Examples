Public MustInherit Class Vehicle
    ' NOTE: This is an Abstract Class that has not methods
    '       It can only be used by inheriting these properties.
    '       into the Car or Truck Class

    Public Property Make As String
    Public Property Model As String
    Public Property Year As Integer
    Public Property Colour As String

    Public MustOverride Function FormatMe() As String

End Class
