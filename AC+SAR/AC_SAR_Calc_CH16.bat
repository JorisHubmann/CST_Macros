
' genererates 16 AC Combine Results each for one channel (mode) and the CP Mode
' Calculates SAR for those channels and for CP

Sub Main ()

	Dim Nports,Freq As Integer
	Freq=298
	Nports=16

	Phase = Array(0, 22.5, 45, 67.5, 90, 112.5, 135, 157.5, 180, 202.5, 225, 247.5, 270, 292.5, 315, 337.5)

	Amp = Array(1, 1 ,1 , 1, 1, 1, 1, 1 ,1 ,1 ,1 ,1 ,1 ,1 ,1 ,1)

	Call createModes(Nports,Freq,Phase,Amp)
	Call CalcSAR(Nports,Freq)
End Sub

'%%%%%%%%%%%%%%%


Sub createModes(Nports,Freq,Phase,Amp)
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
	.ResetCombineMonitorFilters
	.AddCombineMonitorFilter("loss (f=" & Freq & ")")
	.AddCombineMonitorFilter("h-field (f=" & Freq & ")")
	.AddCombineMonitorFilter("e-field (f=" & Freq & ")")
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
			.ResetCombineMonitorFilters
			.AddCombineMonitorFilter("loss (f=" & Freq & ")")
			.AddCombineMonitorFilter("h-field (f=" & Freq & ")")
			.AddCombineMonitorFilter("e-field (f=" & Freq & ")")
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

'%%%%%%%%%%%%%%%%%%%%%%

Sub CalcSAR(Nports,Freq)
	Dim i As Double
	Dim objN As String
	Dim averagingMethod As String

	averagingMethod="IEEE/IEC 62704-1"

	'CP-Mode
	objN = "SAR_CP"
	If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor("loss (f=" & Freq & ") [CP]")
			.AverageWeight(10)
			.SetLabel("SAR_CP")
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
		End With

	End If
	
	For i=1 To Nports
		objN = "SAR_Ch" & i
		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
			With SAR
				.reset
				.PowerlossMonitor("loss (f=" & Freq & ") [CH_" & i & "]")
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