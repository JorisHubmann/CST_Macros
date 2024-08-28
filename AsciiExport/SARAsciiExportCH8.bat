' SARAsciiExportCH8 Sleeve

Sub Main ()
	Dim Nports,Freq As Integer
	Dim AName, pDimension, Orientation, Field As String

	Nports=8
	Freq=298
	Stepsize=5

	AName="Sleeve"
	pDimension="2D"
	Field="SAR"
	Orientation="xy"
	
	Call mASCIIExport(Nports,Freq,Stepsize, AName, pDimension, Field, Orientation)

End Sub

Sub mASCIIExport(Nports,Freq,Stepsize, AName, pDimension, Field, Orientation)

	Dim ExportFolder As String
	ExportFolder= "Z:\CST_Daten\Ergebnisse\DumpData"
	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\SAR\SAR_CP") Then
		With ScalarPlot2D
   			.Type ("contour")
   			.PlaneNormal ("z")
  		  	.PlotAmplitude (False)
    		.PlaneCoordinate (0.0)
   			.Quality (60)
		End With

		Plot3DPlotsOn2DPlane True
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_" + Field + "_" + Orientation + ".txt")
			.Mode ("FixedWidth")
			.SetfileType("ascii")
			'.SetSubvolume(-100,100,-100,100,-0,0)
			.UseSubvolume(False)
			.Step(Stepsize)
			.Execute
		End With
	End If

	'Channels
	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\SAR\SAR_Ch" & i) Then
			Plot3DPlotsOn2DPlane True
			Wait(1)
			With ScalarPlot2D
	   			.Type ("contour")
	   			.PlaneNormal ("z")
	  		  	.PlotAmplitude (False)
	    		.PlaneCoordinate (0.0)
	   			.Quality (60)
			End With
			With ASCIIExport
				.Reset
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_" + Field + "_" + "CH_" + i + "_" + Orientation + ".txt")
				.Mode ("FixedWidth")
				.SetfileType("ascii")
				.UseSubvolume(False)
				.Step(Stepsize)
				.Execute
			End With
		End If

	Next

End Sub
