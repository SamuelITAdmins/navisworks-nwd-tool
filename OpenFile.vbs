'Grab the nwf source file path from the arguments
Dim Args, SourceFile, DestinationFile
Set Args = WScript.Arguments
If Args.Count != 2 Then
    WScript.Echo "Error: Source and Destination files not provided!"
    WScript.Quit
End If
SourceFile = args(0)
DestinationFile = args(1)

'DestinationFile = "C:\Navisworks\#####-OverallModel.nwf"
'SourceFile = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"
'Const DestinationFile = "C:\Navisworks\24317-OverallModel.nwf"

'Ensure the Navisworks directory exists
Const strFolder = "C:\Navisworks\"
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(strFolder) Then
    fso.CreateFolder strFolder
End If

'Check to see if the file already exists in the destination folder
If fso.FileExists(DestinationFile) Then
    'Check to see if the file is read-only
    If Not fso.GetFile(DestinationFile).Attributes And 1 Then 
        'The file exists and is not read-only.  Safe to replace the file.
        fso.CopyFile SourceFile, DestinationFile, True
    Else 
        'The file exists and is read-only.
        'Remove the read-only attribute
        fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes - 1
        'Replace the file
        fso.CopyFile SourceFile, DestinationFile, True
        'Reapply the read-only attribute
        fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
    End If
Else
    'The file does not exist in the destination folder.  Safe to copy file to this folder.
    fso.CopyFile SourceFile, DestinationFile, True
End If
Set fso = Nothing

CreateObject("WScript.Shell").Run(DestinationFile)