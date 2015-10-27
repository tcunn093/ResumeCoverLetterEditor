Sub main()

Dim fs: set fs = CreateObject("Scripting.FileSystemObject")
Dim CurrentDirectory, UpdatedFolder, CoverLetter, ResumeFile, Employer, Position


CurrentDirectory = fs.GetAbsolutePathName(".")

UpdatedFolderName = getUpdatedPathName(CurrentDirectory)

UpdatedFolder = CurrentDirectory + "\" + UpdatedFolderName

CoverLetter = getPathFromTypeofDoc("Cover Letter")
ResumeFile = getPathFromTypeofDoc("Resume")

CoverLetterName = fs.GetFileName(CoverLetter)
ResumeFileName = fs.GetFileName(ResumeFile)

Employer = getEmployer()
Position = getPosition()

Call AddEmployerName(UpdatedFolder, CoverLetterName, Employer, CoverLetter, Position)
Call AddEmployerName(UpdatedFolder, ResumeFileName, Employer, ResumeFile, Position)

fs.GetFolder(UpdatedFolder).Name = "Updated for " + Employer + " - " + Position

msgbox ("Complete")

Set fs = Nothing

End Sub


Sub AddEmployerName(UpdatedF, fileName, EmployerName, file, Posit)

Const wdReplaceAll  = 2

Dim objWord, objDoc, objSelection

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Open(file)
Set objSelection = objWord.Selection


objSelection.Find.Text = "#####"
objSelection.Find.Forward = TRUE
objSelection.Find.MatchWholeWord = TRUE

objSelection.Find.Replacement.Text = EmployerName
objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll

objSelection.Find.Text = "$$$$$"
objSelection.Find.Forward = TRUE
objSelection.Find.MatchWholeWord = TRUE

objSelection.Find.Replacement.Text = Posit
objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll

objDoc.SaveAs UpdatedF + "\" + fileName
objDoc.Close

objWord.Quit

Set objSelection = Nothing
Set objDoc = Nothing
Set objWord = Nothing

End Sub


Function getEmployer()

getEmployer = InputBox("Enter the Employer Name")

End Function


Function getPosition()

getPosition = InputBox("Enter the position for which you are applying:")

End Function


Function getUpdatedPathName(Current)

Dim f1

Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Dim f: Set f = fso.GetFolder(Current)

Dim fc: Set fc = f.Subfolders


For Each f1 in fc
	
	if Instr(f1.name, "Updated") then
		getUpdatedPathName = f1.name
	End if

Next

if getUpdatedPathName = "" then
	fso.CreateFolder(Current + "\Updated")
End if

End Function

Call main



Function getPathFromTypeOfDoc(typeOfDoc)

Dim fsob: set fsob = CreateObject("Scripting.FileSystemObject")
Dim absPath: abspath = fsob.GetAbsolutePathName(".")

For Each file In fsob.GetFolder(absPath).Files


	if Instr(fsob.GetExtensionName(file.Name), "doc") then
		
		if Instr(file.name, typeOfDoc) then
			getPathFromTypeOfDoc = file.path
		End if

	End if

Next

if getPathFromTypeOfDoc = "" then

	msgbox "Please place your " + typeOfDoc + " (Microsoft Word file) into the following folder: " + vbCrLf + vbCrLf + absPath + vbCrLf + vbCrLf + "Change any instance of an employer to ##### and any instance of the job position to $$$$$"
	Wscript.Quit	

End if

End Function
