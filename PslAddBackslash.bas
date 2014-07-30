'Add the backslash back to translated string into the javascript

Option Explicit
Sub Main

  'The " mark
  Dim Quotation As String
  Quotation = Chr$(34)

  'The \" mark
  Dim repString1 As String
  repString1 = Chr$(92)+Chr$(34)

  'The ' mark
  Dim Apostrophe As String
  Apostrophe = Chr$(39)

  'The \' mark
  Dim repString2 As String
  repString2 = Chr$(92)+Chr$(39)

  Dim trn As PslTransList
  Dim i As Long

  ' Get Passolo Project
  Dim prj As PslProject
  Set prj = PSL.ActiveProject

  ' Check whether we have open a project or not
  If prj Is Nothing Then
    MsgBox("NO active Passolo project.")
    Exit Sub
  End If

  For Each trn In prj.TransLists

   For i = 1 To trn.StringCount

   Dim tString As PslTransString

   Dim tempString As String
   tempString = ""
   
   Set tString = trn.String(i)

   If InStr(tString.Text, Quotation) > 0 Then

       'Add the backslash in front of the quotation Mark
       tString.Text = Replace(tString.Text, Quotation, repString1)
       tString.TransList.Save

   'If the translated string contains ', replace it with \'
   ElseIf InStr(tString.Text, Apostrophe) > 0 Then

      'Add the backslash in front of the Apostrophe
       tString.Text = Replace(tString.Text, Apostrophe,repString2)
       tString.TransList.Save

   End If

   Next i

  Next trn

 MsgBox("Done!")

End Sub
