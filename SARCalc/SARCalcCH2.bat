' TestPCalc for 2 Sleeve
Sub Main ()

	Dim Nports,Freq As Integer
	Freq=298
	Nports=2

	Call CalcSAR(Nports,Freq)

End Sub

Sub CalcSAR(Nports,Freq)
	Dim i,j As Double
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

