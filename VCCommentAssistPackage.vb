Imports Microsoft.VisualBasic
Imports System
Imports EnvDTE
Imports System.Diagnostics
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.ComponentModel.Design
Imports Microsoft.Win32
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop
Imports Microsoft.VisualStudio.OLE.Interop
Imports Microsoft.VisualStudio.Shell

''' <summary>
''' This is the class that implements the package exposed by this assembly.
'''
''' The minimum requirement for a class to be considered a valid package for Visual Studio
''' is to implement the IVsPackage interface and register itself with the shell.
''' This package uses the helper classes defined inside the Managed Package Framework (MPF)
''' to do it: it derives from the Package class that provides the implementation of the 
''' IVsPackage interface and uses the registration attributes defined in the framework to 
''' register itself and its components with the shell.
''' </summary>
' The PackageRegistration attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class
' is a package.
'
' The InstalledProductRegistration attribute is used to register the information needed to show this package
' in the Help/About dialog of Visual Studio.
    '
' The ProvideMenuResource attribute is needed to let the shell know that this package exposes some menus.

<PackageRegistration(UseManagedResourcesOnly:=True), _
InstalledProductRegistration("#110", "#112", "2.0", IconResourceID:=400), _
ProvideMenuResource("Menus.ctmenu", 1), _
Guid(GuidList.guidVCCommentAssistPkgString)> _
Public NotInheritable Class VCCommentAssistPackage

    Inherits Package

    ''' <summary>
    ''' Default constructor of the package.
    ''' Inside this method you can place any initialization code that does not require 
    ''' any Visual Studio service because at this point the package object is created but 
    ''' not sited yet inside Visual Studio environment. The place to do all the other 
    ''' initialization is the Initialize method.
    ''' </summary>
    Public Sub New()
        Debug.WriteLine(String.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", Me.GetType().Name))
    End Sub



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Overridden Package Implementation
#Region "Package Members"

    ''' <summary>
    ''' Initialization of the package; this method is called right after the package is sited, so this is the place
    ''' where you can put all the initialization code that rely on services provided by VisualStudio.
    ''' </summary>
    Protected Overrides Sub Initialize()
        Debug.WriteLine(String.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", Me.GetType().Name))
        MyBase.Initialize()

        ' Add our command handlers for menu (commands must exist in the .vsct file)
        Dim mcs As OleMenuCommandService = TryCast(GetService(GetType(IMenuCommandService)), OleMenuCommandService)
        If Not mcs Is Nothing Then

            ' 创建工具条的所有按钮以及对应处理函数
            Dim menuIDVCFileHead As New CommandID(GuidList.guidVCCommentAssistCmdSet, CInt(PkgCmdIDList.cmdidVCCommentFileHead))
            Dim menuItemVCFileHead As New MenuCommand(New EventHandler(AddressOf MenuItemVCFileHeadCallback), menuIDVCFileHead)
            mcs.AddCommand(menuItemVCFileHead)


            Dim menuIDClass As New CommandID(GuidList.guidVCCommentAssistCmdSet, CInt(PkgCmdIDList.cmdidVCCommentClass))
            Dim menuItemVCClass As New MenuCommand(New EventHandler(AddressOf MenuItemVCClassCallback), menuIDClass)
            mcs.AddCommand(menuItemVCClass)

            ' Create vc comment file head command for the menu item.
            Dim menuIDFunction As New CommandID(GuidList.guidVCCommentAssistCmdSet, CInt(PkgCmdIDList.cmdidVCCommentFunction))
            Dim menuItemVCFunction As New MenuCommand(New EventHandler(AddressOf MenuItemVCFunctionCallback), menuIDFunction)
            mcs.AddCommand(menuItemVCFunction)

            Dim menuIDOneLine As New CommandID(GuidList.guidVCCommentAssistCmdSet, CInt(PkgCmdIDList.cmdidVCCommentOneLine))
            Dim menuItemVCOneLine As New MenuCommand(New EventHandler(AddressOf MenuItemVCOneLineCallback), menuIDOneLine)
            mcs.AddCommand(menuItemVCOneLine)

            Dim menuIDAlign As New CommandID(GuidList.guidVCCommentAssistCmdSet, CInt(PkgCmdIDList.cmdidVCCommentAlign))
            Dim menuItemVCAlign As New MenuCommand(New EventHandler(AddressOf MenuItemVCAlignCallback), menuIDAlign)
            mcs.AddCommand(menuItemVCAlign)

        End If

        Dim readValue = My.Computer.Registry.GetValue(
            "HKEY_CURRENT_USER\Software\VCCommentAssist", "Copyright_String", Nothing)
        If (readValue Is Nothing) Then
            My.Computer.Registry.SetValue(
                "HKEY_CURRENT_USER\Software\VCCommentAssist", "Copyright_String", copyright_str)
        Else
            copyright_str = readValue
        End If

        readValue = My.Computer.Registry.GetValue(
            "HKEY_CURRENT_USER\Software\VCCommentAssist", "Author_Name", Nothing)
        If (readValue Is Nothing) Then
            My.Computer.Registry.SetValue(
                "HKEY_CURRENT_USER\Software\VCCommentAssist", "Author_Name", author_name)
        Else
            author_name = readValue
        End If

        Dim cstyle As Integer

        cstyle = My.Computer.Registry.GetValue(
            "HKEY_CURRENT_USER\Software\VCCommentAssist", "Comment_Style", -1)
        If (cstyle = -1) Then
            cstyle = comment_style
            My.Computer.Registry.SetValue(
                "HKEY_CURRENT_USER\Software\VCCommentAssist", "Comment_Style", cstyle)
        Else
            comment_style = cstyle
        End If

    End Sub


#End Region


    ''' <summary>
    ''' This function is the callback used to execute a command when the a menu item is clicked.
    ''' See the Initialize method to see how the menu item is associated to this function using
    ''' the OleMenuCommandService service and the MenuCommand class.
    ''' </summary>
    Private Sub MenuItemVCFileHeadCallback(ByVal sender As Object, ByVal e As EventArgs)

        Dim dte As DTE = TryCast(Package.GetGlobalService(GetType(DTE)), DTE)
        FileCommentsEn(dte)

    End Sub

    Private Sub MenuItemVCFunctionCallback(ByVal sender As Object, ByVal e As EventArgs)

        Dim dte As DTE = TryCast(Package.GetGlobalService(GetType(DTE)), DTE)
        FunctionCommentEn(dte)

    End Sub

    Private Sub MenuItemVCClassCallback(ByVal sender As Object, ByVal e As EventArgs)

        Dim dte As DTE = TryCast(Package.GetGlobalService(GetType(DTE)), DTE)
        ClassCommentEn(dte)

    End Sub


    Private Sub MenuItemVCOneLineCallback(ByVal sender As Object, ByVal e As EventArgs)
        ' Show a Message Box to prove we were here

        Dim dte As DTE = TryCast(Package.GetGlobalService(GetType(DTE)), DTE)
        CommentOneLine(dte)
    End Sub


    Private Sub MenuItemVCAlignCallback(ByVal sender As Object, ByVal e As EventArgs)

        Dim dte As DTE = TryCast(Package.GetGlobalService(GetType(DTE)), DTE)
        CodeBlockAlign(dte)

    End Sub

    '判断代码的缩进的字符串个数，用于后面的对齐处理
    Private Sub GetIndentationSpace(ByRef code_string As String, ByRef indent_string As String)
        Dim space_counter As Integer = 0
        While (StrComp(GetChar(code_string, space_counter + 1), " ") = 0)
            space_counter = space_counter + 1
        End While
        indent_string = Space(space_counter)
    End Sub
    's注释的格式
    Public Enum CommentStyle As Integer
        '格式/*!
        Gang_Xing_Tan = 1
        '格式///
        Gang_Gang_Gang = 2
        '格式//!
        Gang_Gang_Tan = 3
    End Enum

    '你只需要修改这个定义就OK了。
    Dim copyright_str As String = "Apache License, Version 2.0 FULLSAIL"
    Dim author_name As String = "Sailzeng <sailerzeng@gmail.com>"
    Dim comment_style As CommentStyle = CommentStyle.Gang_Xing_Tan

    '------------------------------------------------------------------------------
    'SUB DESCRIPTION: A Macro To Comment function 
    'Author :Sail(ZengXing)      2001.2.8
    'Pleasse move cursor to function define first line，run this sub, can create one.
    Private Sub FunctionCommentEn(ByVal dte As DTE)
        'DESCRIPTION: Function comment macros(English version，),

        'Judge file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If

        ' Record Function Start Postion
        Dim x_current As Integer = 0
        Dim y_current As Integer = 0

        'Create a new text document.

        'Get a handle to the new document.
        x_current = dte.ActiveDocument.Selection.CurrentLine
        y_current = dte.ActiveDocument.Selection.CurrentColumn

        ' Record File End Postion
        Dim x_endfile As Integer
        Dim y_endfile As Integer
        dte.ActiveDocument.Selection.EndOfDocument()
        x_endfile = dte.ActiveDocument.Selection.CurrentLine
        y_endfile = dte.ActiveDocument.Selection.CurrentColumn
        dte.ActiveDocument.Selection.MoveTo(x_current, y_current)

        ' Find Function Defination
        Dim str_seltxt As String
        Dim minus_bracket_count As Integer = 0
        Dim plus_bracket_count As Integer = 0
        Dim while_do As Boolean = True
        Dim find_start As Integer = 1

        str_seltxt = ""

        'judge whether  counter of ( and ) is equal, find full function name define.
        'may be have simpler way to get function name . function name end is ';' or "{"  or ":"?
        Do
            minus_bracket_count = 0
            plus_bracket_count = 0
            dte.ActiveDocument.Selection.SelectLine()
            str_seltxt = str_seltxt + dte.ActiveDocument.Selection.Text

            find_start = 1
            While (InStr(find_start, str_seltxt, "(") <> 0)
                find_start = InStr(find_start, str_seltxt, "(") + 1
                plus_bracket_count = plus_bracket_count + 1
            End While
            find_start = 1
            While (InStr(find_start, str_seltxt, ")") <> 0)
                find_start = InStr(find_start, str_seltxt, ")") + 1
                minus_bracket_count = minus_bracket_count + 1
            End While

            'Find full function name string to analysis
            If (plus_bracket_count = minus_bracket_count And plus_bracket_count <> 0) Then
                while_do = False
                Exit Do
            End If

            If (dte.ActiveDocument.Selection.CurrentLine = x_endfile) Then
                MsgBox("Function define error or cursor in error line cannot be found in pairs() ", vbOKOnly, "ERROR")
                Exit Sub
            End If

            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)

        Loop While (while_do = True)
        'Trim SelString And Clean SelString

        '检查<>是否匹配
        Dim minus_angel_count As Integer = 0
        Dim plus_angel_count As Integer = 0
        find_start = 1
        While (InStr(find_start, str_seltxt, "<") <> 0)
            find_start = InStr(find_start, str_seltxt, "<") + 1
            plus_angel_count = plus_angel_count + 1
        End While
        find_start = 1
        While (InStr(find_start, str_seltxt, ">") <> 0)
            find_start = InStr(find_start, str_seltxt, ">") + 1
            minus_angel_count = minus_angel_count + 1
        End While

        If (minus_angel_count <> plus_angel_count) Then
            MsgBox("Function define error or cursor in error line，cannot be found in pairs <>！", vbOKOnly, "ERROR")
        End If


        Dim str_analysis As String
        Dim start_pos As Integer
        Dim str_indent As String = ""

        str_analysis = str_seltxt
        GetIndentationSpace(str_analysis, str_indent)
        start_pos = InStrRev(str_analysis, ")")
        str_analysis = Left(str_analysis, start_pos)
        str_analysis = Replace(str_analysis, Chr(9), " ")
        str_analysis = Replace(str_analysis, Chr(13), " ")
        str_analysis = Replace(str_analysis, Chr(10), " ")
        str_analysis = Trim(str_analysis)

        '将两个空格调调整成一个
        While (InStr(str_analysis, "  ") <> 0)
            str_analysis = Replace(str_analysis, "  ", " ")
        End While
        '替换某些不同的写法，方便我们处理
        str_analysis = Replace(str_analysis, " ***", "*** ")
        str_analysis = Replace(str_analysis, " **", "** ")
        str_analysis = Replace(str_analysis, " *", "* ")
        str_analysis = Replace(str_analysis, " &", "& ")
        str_analysis = Replace(str_analysis, " :", ":")
        str_analysis = Replace(str_analysis, ": ", ":")
        '上面的整理也可能打来两个空格，再次清理
        While (InStr(str_analysis, "  ") <> 0)
            str_analysis = Replace(str_analysis, "  ", " ")
        End While
        str_analysis = Trim(str_analysis)


        str_seltxt = str_analysis
        dte.ActiveDocument.Selection.MoveTo(x_current, 1)
        dte.ActiveDocument.Selection.LineUp()
        dte.ActiveDocument.Selection.EndOfLine()
        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)

        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("/*!")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert(str_indent)
            dte.ActiveDocument.Selection.Insert("* @brief      ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @brief      ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @brief      ")
        End If

        GetTemplateParNameEn(dte, str_analysis, str_indent)
        GetFunctionNameRtnEn(dte, str_analysis, str_indent)
        GetFunctionParEn(dte, str_analysis, str_indent)

        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)

        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("* @note       ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert(str_indent)
            dte.ActiveDocument.Selection.Insert("*/")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @note       ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @note       ")
        End If

    End Sub



    'Judge Windows type is document type and C++ 
    Private Function JudgeWindowsType(ByVal dte As DTE) As Boolean
        ' check if any file is open 
        If dte.ActiveWindow.Type <> EnvDTE.vsWindowType.vsWindowTypeDocument Then
            MsgBox("This macro can only be run when the document window is activated.", vbOKOnly, "ERROR")
            Return False
        End If

        Dim ext_name As String = dte.ActiveWindow.Document.Name
        Dim pos As Integer = InStrRev(ext_name, ".")

        If pos > 0 Then
            ext_name = Right(ext_name, Len(ext_name) - pos + 1)
            ext_name = LCase(ext_name)
        End If
        '注意_是折行，诡异
        If ext_name <> ".h" And ext_name <> ".hpp" And ext_name <> ".hxx" And ext_name <> ".cpp" And ext_name <> ".cc" _
            And ext_name <> ".c" And ext_name <> ".cxx" And ext_name <> ".inc" Then
            MsgBox("This macro can only be run when a document is edited c/c++ file.", vbOKOnly, "ERROR")
            Return False
        End If
        Return True
    End Function


    '用于得到函数名称，返回值
    'A Function To Get Function Name and Function Return Value
    Private Sub GetFunctionNameRtnEn(ByVal dte As DTE, ByRef str_fun As String, ByRef str_indent As String)

        Dim str_fun_return As String = Trim(str_fun)

        Dim function_name As String
        Dim return_name As String
        Dim start_bracket As Integer
        start_bracket = InStr(str_fun_return, "(")
        str_fun_return = Left(str_fun_return, start_bracket - 1)
        str_fun_return = Trim(str_fun_return)

        'Find Blank,如果有返回值，肯定有空格存在
        Dim start_blank As Integer
        'Form Right Find
        start_blank = InStrRev(str_fun_return, " ")

        '构造函数等没有返回值
        If (start_blank = 0) Then
            function_name = str_fun_return
            return_name = "void"
        Else
            return_name = Left(str_fun_return, start_blank)
            Dim nres_size As Integer
            nres_size = Len(str_fun_return) - start_blank
            function_name = Right(str_fun_return, nres_size)
        End If
        function_name = Trim(function_name)
        return_name = Trim(return_name)


        '去掉inline ,空格是避免名字冲突
        If (Left(return_name, 7) = "inline ") Then
            return_name = Right(return_name, Len(return_name) - 7)
            return_name = Trim(return_name)
        End If

        '去掉static,空格是避免名字冲突
        If (Left(return_name, 7) = "static ") Then
            return_name = Right(return_name, Len(return_name) - 7)
            return_name = Trim(return_name)
        End If


        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)
        
        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("* @return     " + return_name)
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @return     " + return_name)
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("/// @return     " + return_name)
        End If

        str_fun = Right(str_fun, Len(str_fun) - start_bracket + 1)

    End Sub

    '类的长注释方式
    Sub ClassCommentEn(ByRef dte As DTE)

        'judage Open windows type or file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If

        Dim str_analysis As String = ""
        Dim str_indent As String = ""
        GetClassDefString(dte, str_analysis)
        GetIndentationSpace(str_analysis, str_indent)

        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)
        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("/*!")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert(str_indent)
            dte.ActiveDocument.Selection.Insert("* @brief      ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @brief      ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @brief      ")
        End If


        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)
        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("*             ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("///             ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//!             ")
        End If


        GetTemplateParNameEn(dte, str_analysis, str_indent)
        ''GetClassNameEn(str_analysis)
        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)

        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("* note        ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert(str_indent)
            dte.ActiveDocument.Selection.Insert("*/")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @note       ")
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @note       ")
        End If
    End Sub
    '得到class 定义的语句
    Private Sub GetClassDefString(ByRef dte As DTE, ByRef str_analysis As String)

        ' 记录开始的位置Record class Start Postion
        Dim x_current As Integer
        Dim y_current As Integer
        x_current = dte.ActiveDocument.Selection.CurrentLine
        y_current = dte.ActiveDocument.Selection.CurrentColumn
        ' Record File End Postion
        Dim x_endfile
        Dim y_endfile
        dte.ActiveDocument.Selection.EndOfDocument()
        x_endfile = dte.ActiveDocument.Selection.CurrentLine
        y_endfile = dte.ActiveDocument.Selection.CurrentColumn
        dte.ActiveDocument.Selection.MoveTo(x_current, y_current)

        ' Find Function Defination
        Dim str_seltxt As String
        Dim while_do As Boolean = True
        str_seltxt = ""

        '找到{，将前面视为类定义
        Do
            dte.ActiveDocument.Selection.SelectLine()
            str_seltxt = str_seltxt + dte.ActiveDocument.Selection.Text

            Dim find_end As Integer = InStr(str_seltxt, "{")
            If find_end <> 0 Then
                while_do = False
            End If

            If (dte.ActiveDocument.Selection.CurrentLine = x_endfile) Then
                MsgBox("Class define error or cursor in error line，cannot be found in pairs()！", vbOKOnly, "ERROR")
                Exit Sub
            End If

            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        Loop While (while_do = True)

        str_analysis = str_seltxt
        Dim start_pos As Integer

        start_pos = InStr(str_analysis, "{")
        str_analysis = Left(str_analysis, start_pos - 1)
        str_analysis = Replace(str_analysis, Chr(9), " ")
        str_analysis = Replace(str_analysis, Chr(13), " ")
        str_analysis = Replace(str_analysis, Chr(10), " ")
        str_analysis = Trim(str_analysis)
        While (InStr(str_analysis, "  ") <> 0)
            str_analysis = Replace(str_analysis, "  ", " ")
        End While

        dte.ActiveDocument.Selection.MoveTo(x_current, 1)
        dte.ActiveDocument.Selection.LineUp()
        dte.ActiveDocument.Selection.EndOfLine()
    End Sub

    '得到类的名字
    Sub GetClassNameEn(ByRef dte As DTE, ByVal str_analysis As String, ByRef str_indent As String)

        str_analysis = Trim(str_analysis)

        Dim colon_pos As Integer = InStr(str_analysis, ":")
        Dim colon_two As Integer = InStr(str_analysis, "::")
        Dim inherit_class As String = ""
        Dim class_name As String = str_analysis

        '自己感觉理论上也不可能出现这个情况
        If colon_pos > 0 And colon_two <> colon_pos Then
            inherit_class = Right(str_analysis, Len(str_analysis) - colon_pos)
            inherit_class = Trim(inherit_class)
            class_name = Left(str_analysis, colon_pos - 1)
            class_name = Trim(class_name)
        End If


        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)
        '如果是类的定义
        If InStr(class_name, "class ") <> 0 Then
            Dim start_pos As Integer = InStr(class_name, "class ")
            class_name = Right(class_name, Len(class_name) - start_pos - 4)
            class_name = Trim(class_name)
            If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                dte.ActiveDocument.Selection.Insert("* @class      " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                dte.ActiveDocument.Selection.Insert("/// @class      " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                dte.ActiveDocument.Selection.Insert("//! @class      " & class_name)
            End If

        End If

        '如果是结构的定义
        If InStr(class_name, "struct ") <> 0 Then
            Dim start_pos As Integer = InStr(class_name, "struct ")
            class_name = Right(class_name, Len(class_name) - start_pos - 5)
            class_name = Trim(class_name)

            If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                dte.ActiveDocument.Selection.Insert("* @struct     " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                dte.ActiveDocument.Selection.Insert("/// @struct     " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                dte.ActiveDocument.Selection.Insert("//! @struct     " & class_name)
            End If
        End If

        '如果是枚举的定义
        If InStr(class_name, "enum ") <> 0 Then
            Dim start_pos As Integer = InStr(class_name, "enum ")
            class_name = Right(class_name, Len(class_name) - start_pos - 3)
            class_name = Trim(class_name)
            If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                dte.ActiveDocument.Selection.Insert("* @enum       " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                dte.ActiveDocument.Selection.Insert("/// @enum       " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                dte.ActiveDocument.Selection.Insert("//! @enum       " & class_name)
            End If
        End If

        '如果是联合的定义
        If InStr(class_name, "union ") <> 0 Then
            Dim start_pos As Integer = InStr(class_name, "union ")
            class_name = Right(class_name, Len(class_name) - start_pos - 4)
            class_name = Trim(class_name)
            If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                dte.ActiveDocument.Selection.Insert("* @union      " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                dte.ActiveDocument.Selection.Insert("/// @union      " & class_name)
            ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                dte.ActiveDocument.Selection.Insert("//! @union      " & class_name)
            End If
        End If

        '
        If Len(inherit_class) = 0 Then
            Exit Sub
        End If

        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert(str_indent)
        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("* @ref        " & inherit_class)
        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @ref        " & inherit_class)
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @ref        " & inherit_class)
        End If
    End Sub

    '用于得到函数参数
    'A Function To Get Function parameter
    Private Sub GetFunctionParEn(ByVal dte As DTE, ByVal str_fun_param As String, ByRef str_indent As String)
        '参数的起始位置
        Dim start_para As Integer = 0
        Dim end_para As Integer = 0

        '得到参数的列表，去除头尾的括号
        start_para = InStr(str_fun_param, "(")
        start_para = Len(str_fun_param) - start_para
        str_fun_param = Right(str_fun_param, start_para)
        str_fun_param = Trim(str_fun_param)
        end_para = InStrRev(str_fun_param, ")") - 1

        If (end_para > 0) Then
            str_fun_param = Left(str_fun_param, end_para)
            str_fun_param = Trim(str_fun_param)
        Else
            str_fun_param = ""
        End If

        If (str_fun_param = "" Or str_fun_param = "void") Then
        Else
            GetParamterNameEn(dte, str_fun_param, str_indent, True)
        End If
    End Sub

    '生产模板参数
    '注意参数用的应用
    Private Sub GetTemplateParNameEn(ByVal dte As DTE, ByRef str_template As String, ByRef str_indent As String)

        Dim str_analysis As String = Trim(str_template)

        '检索关键字，空格用于防止冲突
        Dim start_pos As Integer = InStr(str_analysis, "template")
        '如果不是最开始是
        If start_pos <> 1 Then
            Exit Sub
        End If

        Dim end_pos As Integer = InStr(str_analysis, "(")
        If end_pos = 0 Then
            end_pos = InStrRev(str_analysis, "class ")
            If end_pos = 0 Then
                end_pos = InStrRev(str_analysis, "struct ")
            End If
        End If

        '取得函数的参数等，
        If end_pos > 0 Then
            end_pos = InStrRev(str_analysis, ">", end_pos)
        End If

        '去掉templace等，返回
        If end_pos > 0 Then

            str_template = Right(str_analysis, Len(str_analysis) - end_pos)
            str_analysis = Left(str_analysis, end_pos)
        Else
            Exit Sub
            str_template = ""
        End If

        Dim start_angle As Integer = InStr(str_analysis, "<")
        Dim end_angle As Integer = InStrRev(str_analysis, ">")
        '没有模版标识的符号< >
        If start_angle = 0 Or end_angle = 0 Then
            Exit Sub
        End If
        '去掉<>
        str_analysis = Right(str_analysis, Len(str_analysis) - start_angle)
        str_analysis = Left(str_analysis, end_angle - start_angle - 1)
        str_analysis = Trim(str_analysis)

        '函数不正确
        If Len(str_analysis) = 0 Then
            Exit Sub
        End If

        GetParamterNameEn(dte, str_analysis, str_indent, False)

    End Sub
    '
    Private Sub CheckPlusAndMinusMatching(ByVal str_check As String, ByVal plus_str As String, ByVal minux_str As String, ByRef matching As Boolean)

        matching = False

        Dim minus_count As Integer = 0
        Dim plus_count As Integer = 0

        Dim find_start As Integer = 1

        find_start = 1
        While (InStr(find_start, str_check, plus_str) <> 0)
            find_start = InStr(find_start, str_check, plus_str) + 1
            plus_count = plus_count + 1
        End While
        find_start = 1
        While (InStr(find_start, str_check, minux_str) <> 0)
            find_start = InStr(find_start, str_check, minux_str) + 1
            minus_count = minus_count + 1
        End While

        If plus_count = minus_count Then
            matching = True
        Else
            matching = False
        End If
    End Sub
    '
    '分析得到参数的名称，同时打印出来
    'if_func 为True标识是函数参数，False标识是模版参数
    Private Sub GetParamterNameEn(ByVal dte As DTE, ByVal str_analysis_para As String, ByRef str_indent As String, _
                                  ByVal if_func As Boolean)

        '参数个数
        Dim param_counter As Integer = 0
        Dim str_param As String = ""

        Do
            param_counter = param_counter + 1

            Dim start_para As Integer = 1
            Dim end_para As Integer = 0
            Dim while_do As Boolean = True

            While while_do

                end_para = InStr(start_para, str_analysis_para, ",") - 1

                If (end_para > 0) Then

                    str_param = Left(str_analysis_para, end_para)

                    start_para = end_para + 2

                    Dim matching As Boolean = False
                    CheckPlusAndMinusMatching(str_param, "<", ">", matching)
                    If matching = False Then
                        Continue While
                    End If

                    CheckPlusAndMinusMatching(str_param, "(", ")", matching)
                    If matching = False Then
                        Continue While
                    End If

                    str_analysis_para = Right(str_analysis_para, Len(str_analysis_para) - end_para - 1)
                    str_analysis_para = Trim(str_analysis_para)
                Else
                    str_param = str_analysis_para
                    str_analysis_para = ""
                End If

                while_do = False
            End While

            str_param = Trim(str_param)

            ''doxygen需要参数的名字，要去掉类型，默认参数，最后结尾的[]等

            ''去掉默认参数
            Dim defualt_parm As Integer = InStr(str_param, "=")
            If defualt_parm <> 0 Then
                str_param = Left(str_param, defualt_parm - 1)
                str_param = Trim(str_param)
            End If

            ''去掉[]
            If Right(str_param, 1) = "]" Then
                Dim array_end As Integer = InStrRev(str_param, "[")
                str_param = Left(str_param, array_end - 1)
                str_param = Trim(str_param)
            End If

            If Right(str_param, 1) = ")" Then
                Dim fun_end As Integer = InStrRev(str_param, "(")
                str_param = Left(str_param, fun_end - 1)
                str_param = Trim(str_param)
            End If

            ''得到名字
            Dim name_start As Integer = InStrRev(str_param, " ")
            If name_start <> 0 Then
                str_param = Right(str_param, Len(str_param) - name_start)
            End If


            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert(str_indent)
            If (if_func) Then
                If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                    dte.ActiveDocument.Selection.Insert("* @param      ")
                ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                    dte.ActiveDocument.Selection.Insert("/// @param      ")
                ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                    dte.ActiveDocument.Selection.Insert("//! @param      ")
                End If
            Else
                If (comment_style = CommentStyle.Gang_Xing_Tan) Then
                    dte.ActiveDocument.Selection.Insert("* @tparam     ")
                ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
                    dte.ActiveDocument.Selection.Insert("/// @tparam     ")
                ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
                    dte.ActiveDocument.Selection.Insert("//! @tparam     ")
                End If
            End If
            dte.ActiveDocument.Selection.Text = str_param

        Loop While (Len(str_analysis_para) > 0)
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------
    'SUB DESCRIPTION: A macro to give file English comments
    'Author :Sail (ZengXing)       2001.2.9
    '用于文件头的注释

    Sub FileCommentsEn(ByRef dte As DTE)
        'DESCRIPTION: 文件头的注释宏(英文版，提供英文注释)，用于自动生成文件头的注释。
        ' check if any file is open 

        'judage Open windows type or file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If

        'dte.ActiveDocument.save
        Dim file_name As String = dte.ActiveDocument.Name

        'macro starts here
        'Input Comments in this Macro
        dte.ActiveDocument.Selection.StartOfDocument()
        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        If (comment_style = CommentStyle.Gang_Xing_Tan) Then
            dte.ActiveDocument.Selection.Insert("/*!")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @copyright  2004-" & System.DateTime.Now.Year & _
                                                "  " & copyright_str)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @filename   " & file_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @author     " & author_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @version    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @date       " & System.DateTime.Now.ToLongDateString())
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @brief      ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @details    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("* @note       ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("*/")
            dte.ActiveDocument.Selection.NewLine()

        ElseIf (comment_style = CommentStyle.Gang_Gang_Gang) Then
            dte.ActiveDocument.Selection.Insert("/// @copyright  2004-" & System.DateTime.Now.Year & _
                                                "  " & copyright_str)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @filename   " & file_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @author     " & author_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @version    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @date       " & System.DateTime.Now.ToLongDateString())
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @brief      ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @details    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("/// @note       ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("///             ")
            dte.ActiveDocument.Selection.NewLine()
        ElseIf (comment_style = CommentStyle.Gang_Gang_Tan) Then
            dte.ActiveDocument.Selection.Insert("//! @copyright  2004-" & System.DateTime.Now.Year & _
                                                "  " & copyright_str)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @filename   " & file_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @author     " & author_name)
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @version    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @date       " & System.DateTime.Now.ToLongDateString())
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @brief      ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @details    ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//! @note       ")
            dte.ActiveDocument.Selection.NewLine()
            dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
            dte.ActiveDocument.Selection.Insert("//!             ")
            dte.ActiveDocument.Selection.NewLine()
        End If

        '保存感觉有点过度了
        'dte.ActiveDocument.save
        'End Recording
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------
    'SUB DESCRIPTION: 一个用于将代码块对齐的宏，用于将变量定义，变量赋值语句格式化成同样对齐的格式。
    'Author :Sail (ZengXing)       2011.5.27
    '比如这样的格式 
    'int a;
    'uint32_t b;
    'unsigned int c;
    '格式化后会成为
    'int            a;
    'uint64_t       b;
    'unsigned int   c;
    '操作方法为要进行对齐的对面全部选择后，运行脚本

    Sub CodeBlockAlign(ByRef dte As DTE)

        'judage Open windows type or file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If

        '选择你要进行规整的代码行，多行一起选择。
        Dim line_start As Integer
        Dim line_end As Integer

        line_end = dte.ActiveDocument.Selection.AnchorPoint.Line
        line_start = dte.ActiveDocument.Selection.ActivePoint.Line

        Dim str_seltxt As String
        Dim max_oneblock_len As Integer = 0

        Dim assignment_str As Boolean = False
        Dim definition_str As Boolean = False

        '找到最长的那行
        For line_process = line_start To line_end
            dte.ActiveDocument.Selection.GotoLine(line_process, True)
            str_seltxt = dte.ActiveDocument.Selection.Text

            '如果不是代码，跳过
            Dim two_blcok_end = InStr(1, str_seltxt, ";")
            If (two_blcok_end = 0) Then
                two_blcok_end = InStr(1, str_seltxt, ",")
                If (two_blcok_end = 0) Then
                    Continue For
                End If
            End If

            Dim find_char = InStr(1, str_seltxt, "=")
            If (find_char <> 0) Then
                assignment_str = True
            Else
                find_char = InStr(1, str_seltxt, " ")
                If (find_char <> 0) Then
                    definition_str = True
                Else
                    Continue For
                End If
            End If

            '如果既有赋值语句又有定义语句，我无法帮你完成格式化
            If (assignment_str = True And definition_str = True) Then
                MsgBox("请选择相同的语句进行规整，或者是变量定义语句，或者是变量赋值语句。")
                Exit Sub
            End If

            Dim procss_chr As String
            Dim procss_num As Integer = two_blcok_end
            While (procss_num > 1)
                '逆向找到空格或者=
                procss_chr = Mid(str_seltxt, procss_num, 1)
                If ((assignment_str = True And procss_chr = "=") Or (definition_str = True And procss_chr = " ")) Then

                    Dim one_block_string
                    one_block_string = Mid(str_seltxt, 1, procss_num - 1)
                    one_block_string = RTrim(one_block_string)

                    '找到最长的文本
                    If max_oneblock_len < Len(one_block_string) Then
                        max_oneblock_len = Len(one_block_string)
                    End If

                    Exit While
                End If
                procss_num = procss_num - 1
            End While
        Next

        '没得搞就退出
        If max_oneblock_len = 0 Then
            Exit Sub
        End If

        '多留几个空格表示美观
        max_oneblock_len = max_oneblock_len + 1

        'real to fucking code
        For line_process = line_start To line_end
            dte.ActiveDocument.Selection.GotoLine(line_process, True)
            str_seltxt = dte.ActiveDocument.Selection.Text

            '如果不是代码，跳过
            Dim two_blcok_end = InStr(1, str_seltxt, ";")
            If (two_blcok_end = 0) Then
                two_blcok_end = InStr(1, str_seltxt, ",")
                If (two_blcok_end = 0) Then
                    Continue For
                End If
            End If

            '看到底是定义语句还是是赋值语句，两者处理方法不太一样
            assignment_str = False
            definition_str = False

            Dim find_char = InStr(1, str_seltxt, "=")
            If (find_char <> 0) Then
                assignment_str = True
            Else
                find_char = InStr(1, str_seltxt, " ")
                If (find_char <> 0) Then
                    definition_str = True
                Else
                    Exit Sub
                End If
            End If

            Dim procss_chr As String
            Dim procss_num As Integer = two_blcok_end
            While (procss_num > 1)

                '逆向找到空格或者=，逆向是为了避免干扰
                procss_chr = Mid(str_seltxt, procss_num, 1)

                If ((assignment_str = True And procss_chr = "=") Or (definition_str = True And procss_chr = " ")) Then
                    Dim one_block_string As String
                    Dim two_block_string As String
                    one_block_string = Mid(str_seltxt, 1, procss_num - 1)
                    one_block_string = RTrim(one_block_string)

                    two_block_string = Mid(str_seltxt, procss_num + 1, Len(str_seltxt))
                    two_block_string = LTrim(two_block_string)
                    dte.ActiveDocument.Selection.Text = one_block_string

                    Dim complete_len As Integer
                    complete_len = 0

                    While (complete_len < (max_oneblock_len - Len(one_block_string)))
                        dte.ActiveDocument.Selection.Text = dte.ActiveDocument.Selection.Text + " "
                        complete_len = complete_len + 1
                    End While
                    '如果是赋值语句
                    If (assignment_str) Then
                        dte.ActiveDocument.Selection.Text = dte.ActiveDocument.Selection.Text + "= "
                    End If

                    dte.ActiveDocument.Selection.Text = dte.ActiveDocument.Selection.Text + two_block_string
                    Exit While
                End If
                procss_num = procss_num - 1
            End While

        Next
    End Sub

    'DESCRIPTION: 将当前文件的Tab键转换为4个空格键
    Sub ConvertTabToBlank(ByRef dte As DTE)

        'judage Open windows type or file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If

        Dim sel_text As String
        sel_text = dte.ActiveDocument.Selection.SelectAll
        sel_text = Replace(sel_text, Chr(9), "    ")
        dte.ActiveDocument.Selection.SelectAll()
        dte.ActiveDocument.Selection.Text = sel_text
    End Sub

    '增加一行注释
    Sub CommentOneLine(ByRef dte As DTE)

        'judage Open windows type or file type.
        If JudgeWindowsType(dte) = False Then
            Exit Sub
        End If
        dte.ActiveDocument.Selection.NewLine()
        dte.ActiveDocument.Selection.MoveTo(dte.ActiveDocument.Selection.CurrentLine, 1)
        dte.ActiveDocument.Selection.Insert("///")
    End Sub



End Class
