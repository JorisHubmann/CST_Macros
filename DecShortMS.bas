
' genererates Nports AC Combine Results each for one channel (mode) and the CP Mode
' Calculates SAR for those channels and for CP

Sub Main ()

	Dim Nports, Freq, Phase, Amp, Stepsize, SubV, SubVSAR, PhaseShim, AmpShim, T As Integer
	Dim AName, pDimension, Orientation, ExportFolder As String

	'Variables
	Nports=2 'Number of Channelss
	Stepsize=2 'Stepsize of the Data export
	AName="Gufi20" 'Change for Result Data Name
	pDimension="3D" 'Keep it to 3D
	Orientation=""	'Only be set for 2D and 1D data Export - Keep at ""
	ExportFolder =  "Y:\CST\ShortMS\Decoupling\Ergebnisse" '"Y:\CST\ShortMS\Decoupling\Ergebnisse"'  'change for directory
	SubV= Array(-176,-64,-226,226,-142,142) ' Gufi
	'SubV= Array(-110,110,-116,130,-230,128) 'Head
	Call PowerExport(AName, ExportFolder)
	Call HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)


End Sub
'%%%%%%%%%%%%%%%


Sub HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)


	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [AC1]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_1" + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [AC2]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_2" + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [AC1+2]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_3" + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

End Sub

Sub PowerExport(AName,ExportFolder)
	If SelectTreeItem ("1D Results\Power\Excitation [AC1]") Then 'Name of the Result Folder
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_Power" + "_1" + ".txt") 'Export File name
			.SetfileType("txt") 'Datatype
			.Execute 'Starts Export
		End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC1]\Loss per Material") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Dielectric" + "_1" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC1]\Loss per Material\Voxel Data") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Voxel" + "_1" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	If SelectTreeItem ("1D Results\Power\Excitation [AC2]") Then 'Name of the Result Folder
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_Power" + "_2" + ".txt") 'Export File name
			.SetfileType("txt") 'Datatype
			.Execute 'Starts Export
		End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC2]\Loss per Material") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Dielectric" + "_2" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC2]\Loss per Material\Voxel Data") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Voxel" + "_2" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
	If SelectTreeItem ("1D Results\Power\Excitation [AC1+2]") Then 'Name of the Result Folder
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_Power" + "_3" + ".txt") 'Export File name
			.SetfileType("txt") 'Datatype
			.Execute 'Starts Export
		End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC1+2]\Loss per Material") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Dielectric" + "_3" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
	If SelectTreeItem ("1D Results\Power\Excitation [AC1+2]\Loss per Material\Voxel Data") Then 'Name of the Result Folder
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (ExportFolder + ".\" + AName + "_Voxel" + "_3" + ".txt") 'Export File name
		.SetfileType("txt") 'Datatype
		.Execute 'Starts Export
	End With
	End If
End Sub
