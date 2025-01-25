Const strFolder = "C:\Navisworks\"
	Const Overwrite = True
DestinationFile = "C:\Navisworks\24317-OverallModel.nwf"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(strFolder) Then
  fso.CreateFolder strFolder
End If
	'DestinationFile = "C:\Navisworks\#####-OverallModel.nwf"
	SourceFile = "\\stor-dn-01\projects\Projects\24317_Electra_CO_EPCM\CAD\Piping\Models\_DesignReview\24317-OverallModel.nwf"


Set fso = CreateObject("Scripting.FileSystemObject")
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
