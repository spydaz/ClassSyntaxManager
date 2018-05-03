Imports System
Imports System.Reflection
Imports System.Windows.Forms
''' <summary>
''' ReadsInfo from Class Type objects
''' </summary>
''' <remarks></remarks>
Public Class TypeReader

    Public Sub New()

    End Sub

    'Main
    'Main
    Public Structure MethodItem

        Public TypeMethod As String
        Public TypeMethodInfo As MethodInfo
        Public TypeName As String

    End Structure

    ''' <summary>
    ''' Gets list of names from objlst.
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public ReadOnly Property MasterList As List(Of String)
        Get
            Return mMasterList
        End Get

    End Property

    Private Shared Sub DisplayMethodInfo(ByVal myArrayMethodInfo() As MethodInfo)
        ' Display information for all methods.
        Dim i As Integer
        For i = 0 To myArrayMethodInfo.Length - 1
            Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
            Console.WriteLine((ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "."))
        Next i
    End Sub

    'DisplayMethodInfo
    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <param name="myArrayMethodInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetMethods(ByVal myArrayMethodInfo() As MethodInfo) As List(Of String)
        Dim Lst As New List(Of String)
        ' Display information for all methods.
        Dim i As Integer
        For i = 0 To myArrayMethodInfo.Length - 1
            Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
            Lst.Add(myMethodInfo.Name)
        Next i
        Return Lst
    End Function

    ' Create a class having two public methods and one protected method.
    Public Shared Function AddSyntax(ByRef Syntax As List(Of String), ByRef MasterSyntax As List(Of String)) As List(Of String)
        Dim NewSyn As New List(Of String)

        For Each term As String In Syntax
            NewSyn.Add(term)
        Next
        For Each term As String In MasterSyntax
            NewSyn.Add(term)
        Next
        Return NewSyn
    End Function
    ''' <summary>
    ''' Returns Grammar for Tools Namespace
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetInternalSyntax() As List(Of String)
        Dim Mlst As New List(Of String)
        Dim Lst As New List(Of String)
        Dim mObjClassLst As List(Of Type) = Addtypes()
        'Public
        For Each Clss As Type In mObjClassLst
            Lst = GetPublic(Clss)
            For Each Detail As String In Lst
                Mlst.Add(LCase(Detail))
            Next
        Next
        'Private
        For Each Clss As Type In mObjClassLst
            Lst = GetPrivate(Clss)
            For Each Detail As String In Lst
                Mlst.Add(LCase(Detail))
            Next
        Next
        Return Mlst
    End Function

    Public Function Addtypes(ByRef mType As Type) As List(Of Type)
        Dim mObjClassLst As New List(Of Type)
        mObjClassLst.Add(GetType(Type))
        Return mObjClassLst
    End Function
    ''' <summary>
    ''' In this function Changes eed to be made to refference in Internal Types
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function Addtypes() As List(Of Type)
        Dim mObjClassLst As New List(Of Type)
        ' mObjClassLst.Add(GetType(NewType))
        Return mObjClassLst
    End Function
    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <param name="MyType"></param>
    ''' <remarks></remarks>
    Public Shared Sub ClassToConsole(ByRef MyType As Type)
        Dim Str As String = ""

        ' Get the public methods.
        Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
        Console.WriteLine((ControlChars.Cr + "The number of public methods is " & myArrayMethodInfo.Length.ToString() & "." & vbNewLine))
        ' Display all the public methods.
        DisplayMethodInfo(myArrayMethodInfo)
        ' Get the nonpublic methods.
        Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
        Console.WriteLine((ControlChars.Cr + "The number of protected methods is " & myArrayMethodInfo1.Length.ToString() & "."))
        ' Display all the nonpublic methods.
        DisplayMethodInfo(myArrayMethodInfo1)
    End Sub

    'Main
    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <param name="MyType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ClassToString(ByRef MyType As Type) As String
        Dim Str As String = ""

        ' Get the public methods.
        Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
        Str += ControlChars.Cr + "The number of public methods is " & myArrayMethodInfo.Length.ToString() & "." & vbNewLine
        ' Display all the public methods.
        Dim i As Integer
        For i = 0 To myArrayMethodInfo.Length - 1
            Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
            Str += ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "." & vbNewLine
        Next i
        ' Get the nonpublic methods.
        Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
        Str += ControlChars.Cr + "The number of protected methods is " & myArrayMethodInfo1.Length.ToString() & "." & vbNewLine
        ' Display all the nonpublic methods.
        Dim j As Integer
        For j = 0 To myArrayMethodInfo1.Length - 1
            Dim myMethodInfo As MethodInfo = CType(myArrayMethodInfo(i), MethodInfo)
            Str += ControlChars.Cr + "The name of the method is " & myMethodInfo.Name & "." & vbNewLine
        Next j
        Return Str
    End Function


    Public Shared Function GetMethodsList(ByRef Types As List(Of Type)) As List(Of MethodItem)
        Dim Lst As New List(Of MethodItem)
        'Dim Types As List(Of Type) = TypeReader.Addtypes
        For Each MyType As Type In Types
            Dim NewItem As New MethodItem
            ' Get the public methods.
            Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or
                                                                      BindingFlags.Instance Or
                                                                      BindingFlags.DeclaredOnly Or
                                                                    BindingFlags.Static))

            Dim i As Integer
            For i = 0 To myArrayMethodInfo.Length - 1
                Dim myMethodInfo As MethodInfo = myArrayMethodInfo(i)
                NewItem.TypeName = MyType.FullName
                NewItem.TypeMethod = myMethodInfo.Name
                NewItem.TypeMethodInfo = myMethodInfo
                Lst.Add(NewItem)
            Next i
        Next
        Return Lst
    End Function

    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <param name="MyType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPrivate(ByRef MyType As Type) As List(Of String)
        Dim ClassGrammar As New List(Of String)

        ' Get the public methods.
        Dim myArrayMethodInfo1 As MethodInfo() = MyType.GetMethods((BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly))
        ClassGrammar = GetMethods(myArrayMethodInfo1)
        Return ClassGrammar
    End Function

    'Main
    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <param name="MyType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPublic(ByRef MyType As Type) As List(Of String)
        Dim ClassGrammar As New List(Of String)
        ' Get the public methods.
        Dim myArrayMethodInfo As MethodInfo() = MyType.GetMethods((BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly Or BindingFlags.Static))
        ClassGrammar = GetMethods(myArrayMethodInfo)
        Return ClassGrammar
    End Function

    Public Shared Sub UpdateTreeViewControl(ByRef TrContol As TreeView)

        Dim Types As List(Of MethodItem) = TypeReader.GetMethodsList(TypeReader.Addtypes)


        Dim Lst As New List(Of String)
        For Each mtype As MethodItem In Types

            If Lst.Contains(mtype.TypeName) Then
            Else
                Lst.Add(mtype.TypeName)
            End If
        Next
        For Each str As String In Lst
            Dim node As New TreeNode
            node.Text = str
            For Each mytype As TypeReader.MethodItem In Types
                If mytype.TypeName = str Then
                    node.Nodes.Add("<" & mytype.TypeMethod & ">" & " Syntax:  " & vbNewLine & mytype.TypeMethodInfo.ToString)

                End If
            Next
            TrContol.Nodes.Add(node)
        Next
    End Sub

    ''' <summary>
    ''' Build list with objClassLst
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BuildMasterList()
        Dim Lst As New List(Of String)
        For Each Clss As Type In ObjClassLst
            Lst = GetPublic(Clss)
            For Each Detail As String In Lst
                mMasterList.Add(Detail)
            Next
        Next
    End Sub

    Private Shared mMasterList As New List(Of String)

    ''' <summary>
    ''' GetType(ClassName(Ref))
    ''' </summary>
    ''' <remarks></remarks>
    Public ObjClassLst As New List(Of Type)

End Class

