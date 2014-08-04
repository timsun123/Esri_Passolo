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

Dim fso, MyFile, FileName

For Each trn In prj.TransLists

    Set fso = CreateObject("Scripting.FileSystemObject")

    Lang = trn.Language.LangCode

    FileName = prj.Location + "\" + prj.Name + "_" + Lang + "_" + trn.Title + ".xliff"

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
