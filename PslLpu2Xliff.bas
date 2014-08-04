'Export the Lpu project file to xliff file

Option Explicit
Sub Main

Dim prj As PslProject
Set prj = PSL.ActiveProject

'Check whether we have open a project or not
If prj Is Nothing Then
	MsgBox("No active Passolo project.")
	Exit Sub
End If

Dim Lang As String
Dim trn As PslTransList
Dim i As Long

Dim fso, fso1, MyFile, FileName, f, f1
Dim Path As String
Dim Path1 As String

Set fso = CreateObject("Scripting.FileSystemObject")

    Path = prj.Location + "\" + prj.Name

    If (fso.FolderExists(Path)) = False Then

    Set f = fso.CreateFolder(Path)

    End If


For Each trn In prj.TransLists

    If StrComp(Lang, trn.Language.LangCode) <> 0 Then

    Lang = trn.Language.LangCode

    Path = prj.Location + "\" + prj.Name + "\" + Lang

      Set fso1 = CreateObject("Scripting.FileSystemObject")

        If (fso1.FolderExists(Path)) = False Then

           Set f1 = fso1.CreateFolder(Path)

      End If

    End If

    FileName = Path + "\" + trn.Title + ".xliff"

    Set MyFile = fso.OpenTextFile(FileName, 2, True)

    MyFile.WriteLine "<?xml version=""1.0"" ?><xliff version=""1.0"">"

    MyFile.WriteLine "  <file original=""global"" source-language=""en_US"" target=""" + Lang +""" datatype=""plaintext"">"

    MyFile.WriteLine "    <body>"

	For i = 1 To trn.StringCount

      Dim tString As PslTransString

      Set tString = trn.String(i)

      MyFile.WriteLine "      <trans-unit id=""" + CStr(tString.Number) + """>"

      MyFile.WriteLine "        <source>" + tString.SourceText + "</source>"

      MyFile.WriteLine "        <target>" + tString.Text + "</target>"

      MyFile.WriteLine "        <note>" + tString.Comment + "</note>"

      MyFile.WriteLine "      </trans-unit>"

	Next i

MyFile.WriteLine "    </body>"
MyFile.WriteLine "  </file></xliff>"

MyFile.Close

Next trn

MsgBox ("Successfully exported Xliff to the same folder which contains the lpu file!")

End Sub

Function CreateFolder



End Function



