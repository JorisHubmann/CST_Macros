' AC_2
' genererates 2 AC Combine Results each for one channel (mode) and the CP Mode

Sub Main ()

	Dim Nports,Freq As Integer
	Freq=298
	Nports=2

	Phase = Array(0, 0)

	Amp = Array(1, 1)

	Call createModes(Nports,Freq,Phase,Amp)

End Sub

Sub createModes(Nports,Freq,Phase,Amp)
Dim i As Double


With SimulationTask
	.reset
	.name("CP")
	.type("AC")
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

'# Individual Chan

For i=1 To Nports
	With SimulationTask
		.reset
		.name("CH_" & i)
		.type("AC")
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