Attribute VB_Name = "safefilenames"
Option Explicit

Public Function SafeFileName(ByVal File As String) As String


Dim FF As String

    FF = Replace$(File, vbNullChar, "_")
    FF = Replace$(FF, Chr$(34), "_")
    FF = Replace$(FF, "\", "_")
    FF = Replace$(FF, "/", "_")
    FF = Replace$(FF, ":", "_")
    FF = Replace$(FF, "*", "_")
    FF = Replace$(FF, "?", "_")
    FF = Replace$(FF, "<", "_")
    FF = Replace$(FF, ">", "_")
    FF = Trim$(Replace$(FF, "|", "_"))
    SafeFileName = FF

End Function

