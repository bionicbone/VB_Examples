Module Methods_Overloaded_Example

    ' The function superSecretFormula is overloaded because there are
    ' two methods with the same name but each accepts different paramters.

    Sub overloadedMethodsExample()
        Debug.WriteLine(superSecretFormula())
        Debug.WriteLine(superSecretFormula("Kevin"))
    End Sub

    Function superSecretFormula() As String
        Return "Hello World"
    End Function

    Function superSecretFormula(ByVal name As String) As String
        Return String.Format("Hello {0}", name)
    End Function

End Module
