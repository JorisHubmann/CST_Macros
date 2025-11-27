'#Language "WWB-COM"

' Calc_Export

Sub Main ()
	Dim Nports,Freq As Integer
	Dim AName, pDimension, Orientation, ExportFolder	As String

	Nports=2
	Stepsize=2
	Freq=297.2

	AName="PA2CH_FR4PLA" 'Change for Result Data Name
	pDimension="3D"
	Orientation=""
	ExportFolder = "Z:\CST_Daten\Simulationen\ElementSimulation\PA\Ergebnisse" 'change for directory

	Phase = Array(0, 0)

	Amp = Array(Sqr(2), Sqr(2))

	'Call createModes(Nports,Freq,Phase,Amp)
	'Call CalcSAR(Nports,Freq)
	'Call SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)
	Call HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)
	'Call EExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)
	'Call SCExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)
End Sub

'%%%%%%%%%%%%%%%%%%%% Create Modes (CP & CH)

Sub createModes(Nports,Freq,Phase,Amp)

	Dim i As Double
	'CO
	With SimulationTask
		.reset
		.name("CP")
		.type("AC")
		If .DoesExist Then
			.delete
		End If
		If Not .DoesExist Then
			.create
			'.ResetCombineMonitorFilters
			'.AddCombineMonitorFilter("loss (f=" + Freq + ")")
			'.AddCombineMonitorFilter("h-field (f=" + Freq + ")")
			'.AddCombineMonitorFilter("e-field (f=" + Freq + ")")
			.SetProperty("docombineresults",1)
			.SetProperty("blocknameforcombineresults","MWSSCHEM1")
			For i=1 To Nports
				.SetComplexPortExcitation(i ,Amp(i-1), Phase(i-1))
				.SetPortSourceType(i, "Signal")
				.SetProperty ("enabled", "True")
			Next
			.Update
		End If
		.setProperty("enabled","False")
	End With

	'# Individual Channels

	For i=1 To Nports
		With SimulationTask
			.reset
			.name("CH_" & i)
			.type("AC")

			If .DoesExist Then
				.delete
			End If
			If Not .DoesExist Then
				.create
				'.ResetCombineMonitorFilters
				'.AddCombineMonitorFilter("loss (f=" + Freq + ")")
				'.AddCombineMonitorFilter("h-field (f=" + Freq + ")")
				'.AddCombineMonitorFilter("e-field (f=" + Freq + ")")
				.SetProperty("docombineresults",1)
				.SetProperty("blocknameforcombineresults","MWSSCHEM1")
				.SetComplexPortExcitation(i, 1, 0)
				.SetPortSourceType(i, "Signal")
				.SetProperty ("enabled", "True")
				.Update
			End If
			.SetProperty ("enabled", "False")
		End With
	Next
End Sub


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% SAR Calc

Sub CalcSAR(Nports,Freq)
	Dim i As Double
	Dim objN As String
	Dim averagingMethod As String

	averagingMethod="IEEE/IEC 62704-1"

	'CP-Mode
	'objN = "SAR_CP"
	'If Resulttree.DoesTreeItemExist("2D/3D Results\SAR\" & objN ) Then
	'	With Resulttree
	'		.Name ("2D/3D Results\SAR\" & objN )
	'		.Delete
	'	End With
	'End If
	'If Not Resulttree.DoesTreeItemExist("2D/3D Results\SAR\" & objN ) Then
	'	With SAR
	'		.reset
	'		.PowerlossMonitor("loss (f=297.2) [CP]")
	'		.AverageWeight(10)
	'		.SetLabel("SAR_CP")

			'# set to 1W stimulated power
	'		.SetOption("scale stimulated")
	'		.SetOption("rescale 1")
	'		.SetOption (averagingMethod)
	'		.Create
	'	End With

'	End If

	For i=1 To Nports
		objN = "SAR_CH" & i
		If Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With Resulttree
				.Name ("2D/3D Results\SAR\" & objN )
				.Delete
			End With
		End If
		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With SAR
				.reset
				.PowerlossMonitor("loss (f=297.2) [CH" & i & "]")
				.AverageWeight(10)
				'.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
				'.SetOption("subvolume only")
				.SetLabel(objN)
				'# set to 1W stimulated power
				.SetOption("scale stimulated")
				.SetOption("rescale 1")
				.SetOption (averagingMethod)
				.Create
			End With
		End If

	Next
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Export

Sub SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)
	Dim i As Integer

	Dim Export As String
	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\SAR\SAR_CH1") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(-110,110,-100,120,-82,182)
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	'For i=1 To Nports
'		If SelectTreeItem ("2D/3D Results\SAR\SAR_CH" & i) Then
		'	Plot3DPlotsOn2DPlane False
		'	Wait(1)
		'	With ASCIIExport
		'		.Reset
		'		.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + "_CH_" & i & ".h5")
		'		.Mode ("FixedWidth")
		'		.SetfileType("hdf5")
		'		.SetSubvolume(-290,-40,-100,100,-100,200) 'Change for ROI
		'		.UseSubvolume(True)
		'		.Step(Stepsize)
		'		.Execute
		'	End With
	'	End If
	'Next
End Sub

Sub HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)


	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CH1]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_CH_1.h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(-110,110,-100,120,-82,182)
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	'For i=1 To Nports
	'	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CH_" & i &"]") Then
'
	'		Plot3DPlotsOn2DPlane False
	'		Wait(1)
	'		With ASCIIExport
	'			.Reset
	'			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_CH_"&i&".h5")
		'		.Mode ("FixedWidth")
		'		.SetfileType("hdf5")
		'		.SetSubvolume(-290,-40,-100,100,-84,100) 'Change for ROI
		'		.UseSubvolume(True)
		'		.Step(Stepsize)
		'		.Execute
		'	End With
	'	End If
	'Next

End Sub

Sub EExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder)
	Dim i As Integer
	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\E-Field\e-field (f=297.2) [CP]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_E" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(-290,-40,-100,100,-100,200) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\E-Field\e-field (f=297.2) [CH_" & i &"]") Then

			Plot3DPlotsOn2DPlane False
			Wait(1)
			With ASCIIExport
				.Reset
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_E" + Orientation + "_CH_" & i &".h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(-290,-40,-100,100,-100,200) 'Change for ROI
				.UseSubvolume(True)
				.Step(Stepsize)
				.Execute
			End With
		End If
	Next
End Sub

Sub SCExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,ExportFolder)


	'MkDir ExportFolder
	'CP-Mode
'	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CP]") Then

	'	Plot3DPlotsOn2DPlane False
	'	Wait(1)
	'	With ASCIIExport
	'		.Reset
	'		.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + ".h5")
	'		.Mode ("FixedWidth")
	'		.SetfileType("hdf5")
	'		.SetSubvolume(-290,-40,-100,100,-100,200) 'Change for ROI
	'		.UseSubvolume(True)
	'		.Step(Stepsize)
	'		.Execute
	'	End With
	'End If

	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\Surface Current\surface current (f=297.2) [CH_" & i &"]") Then

			Plot3DPlotsOn2DPlane False
			Wait(1)
			With ASCIIExport
				.Reset
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SC" + Orientation + "_CH_"&i&".h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(0,0,-150,150,-150,150) 'Change for ROI
				.UseSubvolume(True)
				.Step(Stepsize)
				.Execute
			End With
		End If
	Next

End Sub
