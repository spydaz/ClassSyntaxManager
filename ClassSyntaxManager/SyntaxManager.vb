Imports System.Drawing
Imports System.Windows.Forms
''' <summary>
''' Colors Words in RichText Box
''' </summary>
Public Class SyntaxHighlighter

    Private Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox)
        Dim startindex As Integer = 0
        start = 0
        indexOfSearchText = 0

        If SearchStr <> "" Then

            SearchStr = SearchStr & " "
            If SearchStr.Length > 0 Then
                startindex = FindText(Rtb, SearchStr, start, Rtb.Text.Length)
            End If

            ' If string was found in the RichTextBox, highlight it
            If startindex >= 0 Then
                ' Set the highlight color as red
                Rtb.SelectionColor = Color.CadetBlue

                ' Find the end index. End Index = number of characters in textbox
                Dim endindex As Integer = SearchStr.Length
                ' Highlight the search string
                Rtb.Select(startindex, endindex)

                ' mark the start position after the position of last search string
                start = startindex + endindex

            End If
        Else
        End If
        Rtb.Select(Rtb.TextLength, Rtb.TextLength)
    End Sub

    Private Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox, ByRef MyColor As Color)
        Dim startindex As Integer = 0
        start = 0
        indexOfSearchText = 0

        If SearchStr <> "" Then

            SearchStr = SearchStr & " "
            If SearchStr.Length > 0 Then
                startindex = FindText(Rtb, SearchStr, start, Rtb.Text.Length)
                If startindex = 0 Then
                    startindex = FindText(Rtb, LCase(SearchStr), start, Rtb.Text.Length)
                End If
                If startindex = 0 Then
                    startindex = FindText(Rtb, UCase(SearchStr), start, Rtb.Text.Length)
                End If
            Else
            End If

            ' If string was found in the RichTextBox, highlight it
            If startindex >= 0 Then
                ' Set the highlight color as red
                Rtb.SelectionColor = MyColor

                ' Find the end index. End Index = number of characters in textbox
                Dim endindex As Integer = SearchStr.Length
                ' Highlight the search string
                Rtb.Select(startindex, endindex)

                ' mark the start position after the position of last search string
                start = startindex + endindex

            End If
        Else
        End If
        Rtb.Select(Rtb.TextLength, Rtb.TextLength)
    End Sub

    Private Shared Function FindText(ByRef Rtb As RichTextBox, SearchStr As String,
                        ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer

        ' Unselect the previously searched string
        If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
            Rtb.Undo()
        End If

        ' Set the return value to -1 by default.
        Dim retVal As Integer = -1

        ' A valid starting index should be specified. if indexOfSearchText = -1, the end of search
        If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then

            ' A valid ending index
            If searchEnd > searchStart OrElse searchEnd = -1 Then

                ' Find the position of search string in RichTextBox
                indexOfSearchText = Rtb.Find(SearchStr, searchStart, searchEnd, RichTextBoxFinds.WholeWord)

                ' Determine whether the text was found in richTextBox1.
                If indexOfSearchText <> -1 Then
                    ' Return the index to the specified search text.
                    retVal = indexOfSearchText
                End If

            End If
        End If
        Return retVal

    End Function

    ''' <summary>
    ''' Searches For Internal Syntax
    ''' </summary>
    ''' <param name="Rtb"></param>
    ''' <remarks></remarks>
    Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox)
        'Searches Basic Syntax
        For Each wrd As String In SyntaxTerms
            start = 0
            indexOfSearchText = 0

            ColorSearchTerm(wrd, Rtb)

        Next
    End Sub

    ''' <summary>
    ''' Searches for Specific Word to colorize (Blue)
    ''' </summary>
    ''' <param name="Rtb">      </param>
    ''' <param name="SearchStr"></param>
    ''' <remarks></remarks>
    Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String)

        start = 0
        indexOfSearchText = 0

        ColorSearchTerm(SearchStr, Rtb)

    End Sub

    ''' <summary>
    ''' Searches for Specfic word to colorize specified color
    ''' </summary>
    ''' <param name="Rtb">      </param>
    ''' <param name="SearchStr"></param>
    ''' <param name="MyColor">  </param>
    ''' <remarks></remarks>
    Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String, MyColor As Color)

        start = 0
        indexOfSearchText = 0

        ColorSearchTerm(SearchStr, Rtb, MyColor)

    End Sub

    Private Shared indexOfSearchText As Integer = 0
    Private Shared start As Integer = 0

    Private mGrammar As New List(Of String)

    Public Shared SyntaxTerms() As String = {"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "
ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "
CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "
CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "
DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "
FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "
GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "
LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM", "
LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "
MOVEPREVIOUS", "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING", "
NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "
PRINT", "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "
REFRESH", "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "
RIGHT$", "RMDIR", "RND", "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "
SELECT", "SENDKEYS", "SET", "SETATTR", "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED", "SHELL", "SHOW", "
SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR", "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "
STRING$", "SUB", "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$", "TIMER", "
TIMESERIAL", "TIMEVALUE", "TO", "TRIM", "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "
UNLOAD", "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "
WIDTH", "WRITE", "XOR", "YEAR", "ZORDER"}


    Public Function IsVBKeyword(ByRef keyword As String) As Boolean
        keyword = keyword.ToUpper()
        If keyword = "ADDHANDLER" Or keyword = "ADDRESSOF" Or keyword = "ALIAS" Or
           keyword = "AND" Or keyword = "ANDALSO" Or keyword = "AS" Or keyword = "BYREF" Or
           keyword = "BOOLEAN" Or keyword = "BYTE" Or keyword = "BYVAL" Or keyword = "CALL" Or
           keyword = "CASE" Or keyword = "CATCH" Or keyword = "CBOOL" Or keyword = "CBYTE" Or
           keyword = "CCHAR" Or keyword = "CDATE" Or keyword = "CDEC" Or keyword = "CDBL" Or
           keyword = "CHAR" Or keyword = "CINT" Or keyword = "CLASS" Or keyword = "CLNG" Or
           keyword = "COBJ" Or keyword = "CONST" Or keyword = "CONTINUE" Or keyword = "CSBYTE" Or
           keyword = "CSHORT" Or keyword = "CSNG" Or keyword = "CSTR" Or keyword = "CTYPE" Or
           keyword = "CUINT" Or keyword = "CULNG" Or keyword = "CUSHORT" Or keyword = "DATE" Or
           keyword = "DECIMAL" Or keyword = "DECLARE" Or keyword = "DEFAULT" Or keyword = "DELEGATE" Or
           keyword = "DIM" Or keyword = "DIRECTCAST" Or keyword = "DOUBLE" Or keyword = "DO" Or keyword = "EACH" Or
           keyword = "ELSE" Or keyword = "ELSEIF" Or keyword = "END" Or keyword = "ENDIF" Or keyword = "ENUM" Or
           keyword = "ERASE" Or keyword = "ERROR" Or keyword = "EVENT" Or keyword = "EXIT" Or
           keyword = "FALSE" Or keyword = "FINALLY" Or keyword = "FOR" Or keyword = "FRIEND" Or
           keyword = "FUNCTION" Or keyword = "GET" Or keyword = "GETTYPE" Or keyword = "GLOBAL" Or
           keyword = "GOSUB" Or keyword = "GOTO" Or keyword = "HANDLES" Or keyword = "IF" Or
           keyword = "IMPLEMENTS" Or keyword = "IMPORTS" Or keyword = "IN" Or keyword = "INHERITS" Or
           keyword = "INTEGER" Or keyword = "INTERFACE" Or keyword = "IS" Or keyword = "ISNOT" Or
           keyword = "LET" Or keyword = "LIB" Or keyword = "LIKE" Or keyword = "LONG" Or
           keyword = "LOOP" Or keyword = "ME" Or keyword = "MOD" Or keyword = "MODULE" Or
           keyword = "MUSTINHERIT" Or keyword = "MUSTOVERRIDE" Or keyword = "MYBASE" Or
           keyword = "MYCLASS" Or keyword = "NAMESPACE" Or keyword = "NARROWING" Or
           keyword = "NEW" Or keyword = "NEXT" Or keyword = "NOT" Or keyword = "NOTHING" Or
           keyword = "NOTINHERITABLE" Or keyword = "NOTOVERRIDABLE" Or keyword = "OBJECT" Or
           keyword = "ON" Or keyword = "OF" Or keyword = "OPERATOR" Or keyword = "OPTION" Or
           keyword = "OPTIONAL" Or keyword = "OR" Or keyword = "ORELSE" Or keyword = "OVERLOADS" Or
           keyword = "OVERRIDABLE" Or keyword = "OVERRIDES" Or keyword = "PARAMARRAY" Or keyword = "PARTIAL" Or
           keyword = "PRIVATE" Or keyword = "PROPERTY" Or keyword = "PROTECTED" Or keyword = "PUBLIC" Or
           keyword = "RAISEEVENT" Or keyword = "READONLY" Or keyword = "REDIM" Or keyword = "REM" Or
           keyword = "REMOVEHANDLER" Or keyword = "RESUME" Or keyword = "RETURN" Or keyword = "SBYTE" Or
           keyword = "SELECT" Or keyword = "SET" Or keyword = "SHADOWS" Or keyword = "SHARED" Or
           keyword = "SHORT" Or keyword = "SINGLE" Or keyword = "STATIC" Or keyword = "STEP" Or
           keyword = "STOP" Or keyword = "STRING" Or keyword = "STRUCTURE" Or keyword = "SUB" Or
           keyword = "SYNCLOCK" Or keyword = "THEN" Or keyword = "THROW" Or keyword = "TO" Or
           keyword = "TRUE" Or keyword = "TRY" Or keyword = "TRYCAST" Or keyword = "TYPEOF" Or
           keyword = "WEND" Or keyword = "VARIANT" Or keyword = "UINTEGER" Or keyword = "ULONG" Or
           keyword = "USHORT" Or keyword = "USING" Or keyword = "WHEN" Or keyword = "WHILE" Or
           keyword = "WIDENING" Or keyword = "WITH" Or keyword = "WITHEVENTS" Or keyword = "WRITEONLY" Or
           keyword = "XOR" Or keyword = "#CONST" Or keyword = "#ELSE" Or keyword = "#ELSEIF" Or keyword = "#END" Or keyword = "#IF" = True Then

            Return True
        Else
            Return False
        End If
    End Function

End Class
