Option Strict Off

Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core

Public Class Util
    Public Shared Function GetBuiltInPropertyValue( _
    ByVal objDoc As Object, _
    ByVal PropName As String _
    ) As String

        ' This procedure returns the value of the built-in document
        ' property specified in the strPropName argument for the Office
        ' document object specified in the objDoc argument.

        Dim prpDocProp As DocumentProperty
        Dim varValue As Object

        Const ERR_BADPROPERTY As Long = 5
        Const ERR_BADDOCOBJ As Long = 438
        Const ERR_BADCONTEXT As Long = -2147467259

        Try
            prpDocProp = objDoc.BuiltinDocumentProperties(PropName)

            With prpDocProp
                varValue = .Value
                If Len(varValue) <> 0 Then
                    GetBuiltInPropertyValue = varValue
                Else
                    GetBuiltInPropertyValue = "<Not Set>"
                End If
            End With
        Catch ex As Exception
            Select Case Err.Number
                Case ERR_BADDOCOBJ
                    GetBuiltInPropertyValue = "<No Object.BuiltInProperties>"
                Case ERR_BADPROPERTY
                    GetBuiltInPropertyValue = "<Property not in collection>"
                Case ERR_BADCONTEXT
                    GetBuiltInPropertyValue = "<Value not available in this context>"
                Case Else
                    GetBuiltInPropertyValue = "<BuiltInProperty_Get:Unknown error>"
            End Select
        End Try
    End Function
End Class
