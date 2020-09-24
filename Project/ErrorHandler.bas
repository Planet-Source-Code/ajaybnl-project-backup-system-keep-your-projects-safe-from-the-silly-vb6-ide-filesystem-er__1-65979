Attribute VB_Name = "errHnd"
Option Explicit

Public Function funcHandleError(ByVal sModule As String, _
                                ByVal sProcedure As String, _
                                ByVal oErr As ErrObject, _
                                ByVal Errline As String) As Long


Dim Msg As String

    If Not GetSetting(App.Title, "Errors", "<module-name>-><methord-name>", "NOERR") = "NOERR" Then
        Msg = oErr.Source & " caused error '" & oErr.Description & "' (" & oErr.Number & ")" & vbNewLine & _
         "in module " & sModule & " procedure " & sProcedure & ", line " & Errline & "."
        funcHandleError = MsgBox(Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, "What to you want to do?", oErr.HelpFile, oErr.HelpContext)
    End If

End Function


