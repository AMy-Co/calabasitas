'<declarations>
Public curr_file As Variant
Public curr_dir As Variant
Public output_file As Variant
Public job_number As Variant
Public WO_number As Variant
Public xfdf As Variant
'</declarations>

Const SW_SHOW = 1
Const SW_SHOWMAXIMIZED = 3

Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long



Private Sub ClearStart()
' Deletes everything and pastes Adobe Data
    Dim curr_dir As Variant
    curr_dir = ActiveDocument.Path
    ChangeFileOpenDirectory curr_dir
    
    Dim curr_file As Variant
    curr_file = ActiveDocument.FullName
    
    Documents.Add
    Selection.HomeKey Unit:=wdStory 'Go to the beginning of the document
    Selection.PasteAndFormat (wdFormatPlainText) ' paste
    Selection.WholeStory
    Selection.ClearFormatting
    
End Sub

Private Sub FormatFields()
' Performs majority of formatting
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "LUTRON SERVICES CO., INC."
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
      
    With Selection.Find
        .Text = "7200 SUTER ROAD Schedule Sheet"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "COOPERSBURG PA 18036"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne

    With Selection.Find
        .Text = "Job# :"
        .Replacement.Text = "<field name=""Job Number""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne

    With Selection.Find
        .Text = "Quote# :"
        .Replacement.Text = "^l</value></field>^l<field name=""Quote Number_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
        With Selection.Find
        .Text = "Historic :"
        .Replacement.Text = "^l</value></field>^l<field name=""historicalJN""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Work Order :"
        .Replacement.Text = "^l</value></field>^l<field name=""Work Order#""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
     With Selection.Find
        .Text = "Start Date "
        .Replacement.Text = "^l</value></field>^l<field name=""startDate_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
      With Selection.Find
        .Text = "Lead FSE Information"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Contact :"
        .Replacement.Text = "</value></field>^l<field name=""leadFSEContact_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
       With Selection.Find
        .Text = "Phone :"
        .Replacement.Text = "</value></field>^l<field name=""leadFSEPhone_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
      With Selection.Find
        .Text = "End Date "
        .Replacement.Text = "</value></field>^l<field name=""endDate_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
  
    
         With Selection.Find
        .Text = "Number of Days : "
        .Replacement.Text = "^l</value></field>^l<field name=""numberofDays_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
         
        With Selection.Find
        .Text = "Job Name :"
        .Replacement.Text = "</value></field>^l<field name=""jobName""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
       With Selection.Find
        .Text = "Job Address :"
        .Replacement.Text = "</value></field>^l<field name=""addressStreet""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
      With Selection.Find
        .Text = "City/State :"
        .Replacement.Text = "</value></field>^l<field name=""addressCity""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
        With Selection.Find
        .Text = "Postal Code :"
        .Replacement.Text = "</value></field>^l<field name=""addressZIP""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
          With Selection.Find
        .Text = "Country :"
        .Replacement.Text = "^l</value></field>^l<field name=""addressCountry_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
          With Selection.Find
        .Text = "Region :"
        .Replacement.Text = "^l</value></field>^l<field name=""addressRegion_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
        With Selection.Find
        .Text = "Rep Agency :"
        .Replacement.Text = "</value></field>^l<field name=""repAgency""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
        With Selection.Find
        .Text = "Specifier/Ltg Designer :"
        .Replacement.Text = "^l</value></field>^l<field name=""SPEC-Company""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
        With Selection.Find
        .Text = "Rep contact :"
        .Replacement.Text = "</value></field>^l<field name=""repName""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
          
          
          With Selection.Find
        .Text = "Rep Phone :"
        .Replacement.Text = "</value></field>^l<field name=""repPhone""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
      
       With Selection.Find
        .Text = "[a-zA-Z0-9\-_.]{1,}\@[a-zA-Z0-9\-_.]{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    lutronPM = Selection.Text
    Selection.Delete
      
    With Selection.Find
        .Text = "Project Manager :"
        .Replacement.Text = "^l</value></field>^l<field name=""lutronPM""><value>" + lutronPM + "</value></field>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
      With Selection.Find
        .Text = "Rep Email :"
        .Replacement.Text = "<field name=""repEmail""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Customer/End User Information"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
   
    With Selection.Find
        .Text = "Contact :"
        .Replacement.Text = "</value></field>^l<field name=""CEU-Name""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Email :"
        .Replacement.Text = "^l</value></field>^l<field name=""CEU-Email""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Phone :"
        .Replacement.Text = "</value></field>^l<field name=""CEU-Phone""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Cell Phone :"
        .Replacement.Text = "^l</value></field>^l<field name=""CEU-Phone_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Scheduling Contact Information Job Site Contact Information"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Contact :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Contact_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Contact :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Name""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Company :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Company_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Company :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Company""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Phone :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Phone_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Phone :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Phone""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Fax :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Fax_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Fax :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Fax_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Cell Phone :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Cell_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Cell Phone :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Cell_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Email :"
        .Replacement.Text = "</value></field>^l<field name=""SCHED-Email_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Email :"
        .Replacement.Text = "^l</value></field>^l<field name=""OEU-Email""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Field Engineer(s) :"
        .Replacement.Text = "</value></field>^l<field name=""fseAscName""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Visit Type :"
        .Replacement.Text = "</value></field>^l<field name=""visitType""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "After Hours :"
        .Replacement.Text = "^l</value></field>^l<field name=""afterHours_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Warranty Start Date :"
        .Replacement.Text = "</value></field>^l<field name=""Warranty Start""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "Service Contract :"
        .Replacement.Text = "^l</value></field>^l<field name=""serviceContract_DNE""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = "SystemType :"
        .Replacement.Text = "</value></field>^l<field name=""systemType""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    With Selection.Find
        .Text = ".^pServices"
        .Replacement.Text = "</value></field>^l<field name=""trainingNotes""><value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    
    
    Selection.EndKey Unit:=wdStory 'Go to the end of the document
    Selection.TypeText Text:="</value></field></fields><ids original=""749530BF81975D49A92A5B70875DFA6D"" modified=""8D03AED5F47D4043B8277F3DEDC6C2D1""/></xfdf>"
    
    Selection.HomeKey Unit:=wdStory 'Go to the beginning of the document
    Selection.TypeText Text:="<?xml version=""1.0"" encoding=""UTF-8""?><xfdf xmlns=""http://ns.adobe.com/xfdf/"" xml:space=""preserve""><fields>"
    
    
      
End Sub

Private Sub DeleteSpaces()

    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = "<value> "
        .Replacement.Text = "<value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = " </value>"
        .Replacement.Text = "</value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
Private Sub ReplaceOffenders()

    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = "&"
        .Replacement.Text = "AND"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub


Private Sub DeleteReturns()
' Clears unnecessary returns
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = "^l</value>"
        .Replacement.Text = "</value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = "^p</value>"
        .Replacement.Text = "</value>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub


Private Sub fixResNum()
' extracts the resource number
'
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Find
        .Text = "fseAscName"
        .Forward = True
    End With
    Dim res_num As Variant
    
    Selection.Find.Execute
    Selection.MoveRight Unit:=wdCharacter, Count:=16
    Selection.MoveLeft Unit:=wdCharacter, Count:=6, Extend:=wdExtend
    res_num = Selection.Text
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.TypeText Text:="<field name=""Resource#""><value>"
    Selection.TypeText Text:=res_num
    Selection.TypeText Text:="</value></field>"
End Sub

Private Sub SaveXFDF()
' Performs necessary functions to save data to desktop as JN#_WO#.xfdf


' <----findJN---->
Selection.HomeKey Unit:=wdStory
Selection.WholeStory
With Selection.Find
.Text = "<field name=""Job Number""><value>"
.Forward = True
End With
Dim job_number As Variant 'initialize variable
Selection.Find.Execute
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.MoveRight Unit:=wdCharacter, Count:=7, Extend:=wdExtend
job_number = Selection.Text
'MsgBox job_number
' </----findJN---->

' <----findWO---->
Selection.HomeKey Unit:=wdStory
Selection.WholeStory
With Selection.Find
.Text = "<field name=""Work Order#""><value>"
.Forward = True
End With

Dim WO_number As Variant 'initialize variable
Selection.Find.Execute
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.MoveRight Unit:=wdCharacter, Count:=8, Extend:=wdExtend
WO_number = Selection.Text
'MsgBox WO_number
' </----End findWO---->

' <----Date---->
Dim SO_date As Variant 'initialize variable
SO_date = Format((Year(Now() + 1) Mod 100), _
        "20##") & "-" & _
        Format((Month(Now() + 1) Mod 100), "0#") & "-" & _
        Format((Day(Now()) Mod 100), "0#")
' </----Date---->

' <----Output File Name---->

Dim output_file As Variant  '  initialize variable
output_file = WO_number + "_" + job_number + "_SO_" + SO_date + ".xfdf"

' <----/Output File Name---->

' </----Save xfdf---->


ActiveDocument.SaveAs2 FileName:=output_file, FileFormat:= _
        wdFormatText, LockComments:=False, Password:="", AddToRecentFiles:=False, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
         SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, Encoding:=437, InsertLineBreaks:=False, AllowSubstitutions:=True, _
        LineEnding:=wdCRLF, CompatibilityMode:=0

'Documents.Open FileName:=curr_file
'Documents(curr_dir & "\" & output_file).Close
ActiveDocument.Close

' </----Save xfdf---->

    Dim curr_dir As Variant
    curr_dir = ActiveDocument.Path
    ChangeFileOpenDirectory curr_dir
    
    Dim curr_file As Variant
    curr_file = ActiveDocument.FullName

MsgBox "A file was generated: " & curr_dir & "\" & output_file & ". You may now close this document. After you've loaded data to the Smart Signoff, you can delete the xfdf file.", vbOKOnly, "Hope it worked!"

Dim xfdf As String
xfdf = curr_dir & "\" & output_file

Dim adobe_path As String
adobe_path = "C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe"

ActiveDocument.FollowHyperlink Address:=xfdf

End Sub

Sub PrepareFDF()

' calls everything to output the file
 ClearStart
 FormatFields
 ReplaceOffenders
 DeleteSpaces
 DeleteReturns
 fixResNum
 DeleteSpaces
 DeleteReturns
 Selection.WholeStory
 Selection.ClearFormatting
 SaveXFDF

  
End Sub


