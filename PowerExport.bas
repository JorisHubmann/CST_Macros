'Power Export
Sub Main

	Dim AName, ExportFolder As String

	AName="RD_8" 'Change for Result Data Name
	ExportFolder = "Z:\CST_Daten\Simulationen\ElementSimulation\Ergebnisse\8CH" 'change for directory

	Call PowerExport(AName, ExportFolder)
End Sub


Sub PowerExport(AName,ExportFolder)
	If SelectTreeItem ("1D Results\Power\Excitation [CP]") Then 'Name of the Result Folder
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_Power" + ".txt") 'Export File name
			.SetfileType("txt") 'Datatype
			.Execute 'Starts Export
		End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [CP]\Loss per Material") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Dielectric" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [CP]\Loss per Material\Voxel Data") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Voxel" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
End Sub
