Imports System

Class GuidList
    Private Sub New()
    End Sub

    Public Const guidVCCommentAssistPkgString As String = "9aee591d-45c9-4ca8-8a60-c00e4883093e"
    Public Const guidVCCommentAssistCmdSetString As String = "58f1b9b3-0a7d-4629-b405-761df5fb1650"

    Public Shared ReadOnly guidVCCommentAssistCmdSet As New Guid(guidVCCommentAssistCmdSetString)
End Class