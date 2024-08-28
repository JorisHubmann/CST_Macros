
' genererates 16 AC Combine Results each for one channel (mode) and the CP Mode
' Calculates SAR for those channels and for CP

Sub Main ()

	Dim Nports, Freq, Phase, Amp, Stepsize, SubV, SubVSAR, PhaseShim, AmpShim, T As Integer
	Dim AName, pDimension, Orientation, ExportFolder As String

	'Variables
	Nports=16 'Number of Channels
	Stepsize=1 'Stepsize of the Data export
	AName="LP1x8" 'Change for Result Data Name
	pDimension="3D" 'Keep it to 3D
	Orientation=""	'Only be set for 2D and 1D data Export - Keep at ""
	ExportFolder = "Z:\CST_Daten\Simulationen\ArrayDecision\Results\DP" 'change for directory

	Phase = Array(0, 22.5, 45, 67.5, 90, 112.5, 135, 157.5, 180, 202.5, 225, 247.5, 270, 292.5, 315, 337.5) 'Phase for CP Excitation
	Amp = Array(Sqr(2), Sqr(2) ,Sqr(2) , Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2), Sqr(2)) 'Amplitude for CP Excitation

	SubV= Array(-290,-40,-100,100,-100,200) ' Volume for Export H and E field as well as any other field
	SubVSAR= Array(-318,-40,-234,218,-100,234) ' Volume for SAR Export

	PhaseShim = Array     (-61.9672, 34.0003, 98.664, 174.8744, -72.5277, 35.7277, 133.4349, -125.1603, -93.645, -90.3546, -144.122, 173.0554, -65.6177, 76.0231, 163.2712, -145.6136)
	AmpShim = Array   (4.7275, 4.8906, 4.2786, 4.4227, 6.0144, 5.461, 3.0925, 2.4385, 1.879, 1.1877, 2.0438, 3.4545, 3.6599, 3.7888, 4.1337, 3.8573)
	'Functions
	'Call createModes(Nports,Phase,Amp)
	'Call CalcSAR(Nports)
	'Call SARExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubVSAR)
	'Call HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)
	'Call EExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)
	Call B1Shim(Nports,Stepsize,AName,PhaseShim,AmpShim,pDimension, Orientation, ExportFolder,SubV,SubVSAR)
End Sub

'%%%%%%%%%%%%%%%


Sub createModes(Nports,Phase,Amp)
	Dim i As Double


	With SimulationTask
		.reset
		.name("CP")
		.type("AC")
		If .DoesExist Then
			.delete
		End If
		If Not .DoesExist Then
			.create
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
				.SetProperty("docombineresults",1)
				.SetProperty("blocknameforcombineresults","MWSSCHEM1")
				.SetComplexPortExcitation(i, Sqr(2), 0)
				.SetPortSourceType(i, "Signal")
				.SetProperty ("enabled", "True")
				.Update
			End If
			.SetProperty ("enabled", "False")
		End With
	Next
End Sub

'%%%%%%%%%%%%%%%%%%%%%%

Sub CalcSAR(Nports)
	Dim i As Double
	Dim objN,averagingMethod,T As String


	averagingMethod="IEEE/IEC 62704-1"

	'CP-Mode
	objN = "SAR_CP"
	If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor ("loss (f=297.2) [CP]")
			.AverageWeight(10)
			.SetLabel("SAR_CP")
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
			'MsgBox ("CP max SAR =" & .GetValue("max sar"))
		End With

	End If

	For i=1 To Nports
		objN = "SAR_CH" & i

		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With SAR
				.reset
				.PowerlossMonitor("loss (f=297.2) [CH_" & i & "]")
				.AverageWeight(10)
				.SetLabel(objN)
				'# set to 1W stimulated power
				.SetOption("scale stimulated")
				.SetOption("rescale 1")
				.SetOption (averagingMethod)
				.Create
				'MsgBox ("CH_" & i & " max SAR =" & .GetValue("max sar"))
			End With
		End If
	Next
End Sub

'%%%%%%%%%%%%%%%%%%%%%%%

Sub SARExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubVSAR)

	Dim i As Integer

	'CP-Mode
	If SelectTreeItem ("2D/3D Results\SAR\SAR_CP") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubVSAR(0),SubVSAR(1),SubVSAR(2),SubVSAR(3),SubVSAR(4),SubVSAR(5)) 'Element Selection -290,-40,-100,100,-100,200
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\SAR\SAR_CH" & i) Then
			Plot3DPlotsOn2DPlane False
			Wait(1)
			With ASCIIExport
				.Reset
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + "_CH_" & i & ".h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(SubVSAR(0),SubVSAR(1),SubVSAR(2),SubVSAR(3),SubVSAR(4),SubVSAR(5)) 'Change for ROI
				.UseSubvolume(True)
				.Step(Stepsize)
				.Execute
			End With
		End If
	Next
End Sub

Sub HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)


	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CP]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CH_" & i &"]") Then

			Plot3DPlotsOn2DPlane False
			Wait(1)
			With ASCIIExport
				.Reset
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_CH_" & i & ".h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
				.UseSubvolume(True)
				.Step(Stepsize)
				.Execute
			End With
		End If
	Next

End Sub

Sub EExport(Nports,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)
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
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
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
				.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
				.UseSubvolume(True)
				.Step(Stepsize)
				.Execute
			End With
		End If
	Next
End Sub

'%%%%%%%%%%%%%%

Sub B1Shim(Nports,Stepsize,AName, PhaseShim,AmpShim, pDimension, Orientation, ExportFolder,SubV,SubVSAR)

	Dim i As Double
	Dim objN,averagingMethod,T As String

	averagingMethod="IEEE/IEC 62704-1"

	With SimulationTask
		.reset
		.name("Shim")
		.type("AC")
		If .DoesExist Then
			.delete
		End If
		If Not .DoesExist Then
			.create
			.SetProperty("docombineresults",1)
			.SetProperty("blocknameforcombineresults","MWSSCHEM1")
			For i=1 To Nports
				.SetComplexPortExcitation(i ,AmpShim(i-1), PhaseShim(i-1))
				.SetPortSourceType(i, "Signal")
				.SetProperty ("enabled", "True")
			Next
			.Update
		End If
	.setProperty("enabled","False")
	End With

	'Shim Calc SAR
	objN = "SAR_Shim"
	If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor ("loss (f=297.2) [Shim]")
			.AverageWeight(10)
			.SetLabel(objN)
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
			'MsgBox ("CP max SAR =" & .GetValue("max sar"))
		End With

	End If

	If SelectTreeItem ("2D/3D Results\SAR\SAR_Shim") Then
		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + "_Shim_1.h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubVSAR(1),SubVSAR(2),SubVSAR(3),SubVSAR(4),SubVSAR(5)) 'Element Selection -290,-40,-100,100,-100,200
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [Shim]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + "_Shim_1.h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5)) 'Change for ROI
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

End Sub
