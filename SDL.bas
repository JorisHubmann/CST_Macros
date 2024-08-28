
' genererates 16 AC Combine Results each for one channel (mode) and the CP Mode
' Calculates SAR for those channels and for CP

Sub Main ()

	Dim Nports, Freq, Phase, Amp, Stepsize, SubV, SubVSAR, PhaseShim, AmpShim, T As Integer
	Dim AName, pDimension, Orientation, ExportFolder As String

	'Variables
	Nports=16 'Number of Channels
	Stepsize=2 'Stepsize of the Data export
	AName="SDL_BA" 'Change for Result Data Name
	pDimension="3D" 'Keep it to 3D
	Orientation=""	'Only be set for 2D and 1D data Export - Keep at ""
	ExportFolder = "Z:\CST_Daten\Simulationen\SDL_Paper\Results" 'change for directory

	Phase = Array(0, 45, 90, 135, 180, 225, 270, 315,0, 45, 90, 135, 180, 225, 270, 315) 'Phase for CP Excitation
	'Phase = Array(0, 45, 90, 135, 180, 225, 270, 315, 22.5, 67.5, 112.5, 157.5, 202.5, 247.5, 292.5, 337.5) 'Phase for CP Excitation
	Amp = Array(1/Sqr(8), 1/Sqr(8) ,1/Sqr(8) , 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8), 1/Sqr(8)) 'Amplitude for CP Excitation

	SubV= Array(-270,276,-176,116,-246,126) ' Volume for Export H and E field as well as any other field
	SubVSAR= Array(-270,276,-176,116,-246,126) ' Volume for SAR Export

	'Functions

	Call Spara
	Call Optimizer(Nports)
	Call createModes(Nports,Phase,Amp)
	Call CalcSAR(Nports)
	Call SARExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubVSAR)
	Call HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV)

End Sub

'%%%%%%%%%%%%%%%

Sub Spara
	With SimulationTask
		.Reset
		.Type ("S-Parameters")
		.Name ("Spara")
		If .DoesExist Then
			.delete
		End If
		.Create
		.Update
	End With
End Sub

Sub Optimizer(Nports)

	With SimulationTask
		.Reset
		.Name ("Opt" & Nports)
		If .DoesExist Then
			.delete
		End If
		If Not .DoesExist Then
			.Type ("Optimization")
			.Create
			.Reset
			.Type ("S-Parameters")
			.Name ("Opt" & Nports & "\SPar")
			.SetProperty("maximum frequency range", "True")
			.Create
		End If
	End With

	With DSOptimizer
		.SetSimulationType( "Opt" & Nports)
		.SetOptimizerType("Trust_Region")
		.SetAlwaysStartFromCurrent(True)
		.SetSigma(0.4,"Trust_Region")
		.InitParameterList

		For i=1 To Nports
			.SelectParameter("CM" & i, True)
			.SetParameterInit(5.75)
			.SetParameterMin(5)
			.SetParameterMax(6.5)
			.SelectParameter("Xarm" & i, True)
			.SetParameterInit(328)
			.SetParameterMin(310)
			.SetParameterMax(330)
		Next

'define a goal on the S-Parameters

		Dim gid As Integer
		For i=1 To Nports
			gid = .AddGoal("1DC Primary Result")
			.SelectGoal(gid, True)
			.SetGoal1DCResultName("Opt\SPar\S-Parameters\S" & i & "," & i)
			.SetGoalScalarType("magdB20")
			.SetGoalOperator("<")
			.SetGoalTarget(-30)
			.SetGoalRangeType("single")
			.SetGoalRange(297.2,297.2)
			.SetGoalNormNew ("SumSqDiff")
		Next

'actually start the optimization process
		.Start
	End With

End Sub

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
