
'"C:\Program Files (x86)\CST Studio Suite 2022\AMD64\CST DESIGN ENVIRONMENT_AMD64.exe" -m "C:\Users\Joris\Documents\joris.hu\Meine Bibliotheken\Meine Bibliothek\Promotion\Repository\cst-datahandling\CST_Macros\ServerMacro.bas"

Sub Main ()

	Dim mws,app As Object
	Dim Nports, Freq, Stepsize, SubV, A As Integer
	Dim Phase, Amp, B As Double
	Dim AName, pDimension, Orientation, ExportFolder, ProjectFile, T, C As String



	'Variables
	Nports=2 'Number of Channels
	Stepsize=2 'Stepsize of the Data export
	AName="ServerTest" 'Change for Result Data Name
	pDimension="3D" 'Keep it to 3D
	Orientation=""	'Only be set for 2D and 1D data Export - Keep at ""
	ExportFolder = "Z:\CST_Daten\Simulationen\ServerTest" 'change for directory
	ProjectFile = "Z:\CST_Daten\Simulationen\ServerTest\ServerTest.cst"
	Phase = Array(0, 90) 'Phase for CP Excitation
	Amp = Array(Sqr(2),Sqr(2)) 'Amplitude for CP Excitation total of 1W excitation power
	SubV = Array(-290, -40, -100, 100, -100, 200) ' Volume for Export H and E field as well as any other field for CP

	'Set app = CreateObject ("CSTStudio.Application.2022")
	'Set mws = app.OpenFile (ProjectFile)

	'Call Spara (mws)
	'Call Optimizer(mws)
	'Call createModes(Nports,Phase,Amp,mws)
	'Call HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV,mws)
	'Call SparaExport(ExportFolder)
	'mws.Quit
	T=Array("Spara\S-Parameters", ExportFolder,"50")
	TouchstoneExport ( T)
End Sub

'Spara
Sub Spara(mws)
	With mws.SimulationTask
		.Reset
		.Type ("S-Parameters")
		.Name ("Spara")
		If .DoesExist Then
			.delete
		End If
		.SetProperty("maximum frequency range", "True")
		.Create
		.Update

	End With
End Sub

'Optimizer
Sub Optimizer(mws)
	With mws.SimulationTask
		.Reset
		.Name ("Opt")
		If .DoesExist Then
			.delete
		End If
		If Not .DoesExist Then
			.Type ("Optimization")
			.Create
			.Reset
			.Type ("S-Parameters")
			.Name ("Opt\SPar")
			.SetProperty("maximum frequency range", "True")
			.Create
		End If
	End With

	With mws.DSOptimizer
		.SetSimulationType( "Opt" )
		.SetOptimizerType("Trust_Region")
		.SetAlwaysStartFromCurrent(True)
		.SetSigma(0.4,"Trust_Region")
		.InitParameterList

		.SelectParameter("CM", True)
		.SetParameterInit(10e-12)
		.SetParameterMin(1e-12)
		.SetParameterMax(20e-12)
		.SelectParameter("LD", True)
		.SetParameterInit(40e-9)
		.SetParameterMin(1e-9)
		.SetParameterMax(100e-9)

		.SelectParameter("CM2", True)
		.SetParameterInit(10e-12)
		.SetParameterMin(1e-12)
		.SetParameterMax(20e-12)
		.SelectParameter("LD2", True)
		.SetParameterInit(40e-9)
		.SetParameterMin(1e-9)
		.SetParameterMax(100e-9)
		.DeleteAllGoals

	'define a goal on the S-Parameters

		Dim gid As Integer
		gid = .AddGoal("1DC Primary Result")
		.SelectGoal(gid, True)
		.SetGoal1DCResultName("Opt\SPar\S-Parameters\S1,1")
		.SetGoalScalarType("magdB20")
		.SetGoalOperator("<")
		.SetGoalTarget(-30)
		.SetGoalRangeType("single")
		.SetGoalRange(297.2,297.2)

		gid = .AddGoal("1DC Primary Result")
		.SelectGoal(gid, True)
		.SetGoal1DCResultName("Opt\SPar\S-Parameters\S2,2")
		.SetGoalScalarType("magdB20")
		.SetGoalOperator("<")
		.SetGoalTarget(-30)
		.SetGoalRangeType("single")
		.SetGoalRange(297.2,297.2)

	'actually start the optimization process
		.Start
	End With

End Sub

Sub createModes(Nports,Phase,Amp,mws)
	Dim i As Double

	With mws.SimulationTask
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

End Sub

Sub HExport(Nports,Stepsize, AName, pDimension, Orientation,ExportFolder,SubV,mws)


	'MkDir ExportFolder
	'CP-Mode
	If mws.SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [CP]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With mws.ASCIIExport
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

End Sub



Sub SparaExport(ExportFolder)



End Sub
