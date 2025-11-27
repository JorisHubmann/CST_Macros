
' genererates Nports AC Combine Results each for one channel (mode) and the CP Mode
' Calculates SAR for those channels and for CP

Sub Main ()

	Dim Nports, Freq, Phase, Amp, Stepsize, SubV, SubVSAR, PhaseShim, AmpShim, T As Integer
	Dim AName, pDimension, Orientation, ExportFolder As String

	'Variables
	Nports=8 'Number of Channelss

	'Functions
	Call createModes(Nports)
	Call CalcSAR(Nports)

End Sub

'%%%%%%%%%%%%%%%


Sub createModes(Nports)
	Dim i As Double

'# Individual Channels

	For i=1 To Nports
		With SimulationTask
			.reset
			.name("TxCh" & i)
			.type("AC")
			If .DoesExist Then
				.delete
			End If
			If Not .DoesExist Then
				.create
				.SetProperty("docombineresults",1)
				.SetProperty("blocknameforcombineresults","MWSSCHEM1")
				.SetComplexPortExcitation(i, 1/2, 0)
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

	objN = "TxCh1"

		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With SAR
				.reset
				.PowerlossMonitor("loss (f=297.2) [TxCh1]")
				.AverageWeight(10)
				.SetLabel(objN)
				'# set to 1W stimulated power
				.SetOption("scale stimulated")
				.SetOption("rescale 0.02") ' no rescale für VOP
				.SetOption (averagingMethod)
				.SetOption("no ar results")
				.SetOption("no subvolume")
				.SetOption("opt sarcubes")
				.SetOption("opt longlog")
				.Create
			End With
		End If

	For i=2 To Nports
		objN = "TxCh" & i

		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With SAR
				.reset
				.PowerlossMonitor("loss (f=297.2) [TxCh" & i & "]")
				.AverageWeight(10)
				.SetLabel(objN)
				'# set to 1W stimulated power
				.SetOption("scale stimulated")
				.SetOption("rescale 0.02")
				.SetOption (averagingMethod)
				.SetOption("no ar results")
				.Create
			End With
		End If
	Next
End Sub
