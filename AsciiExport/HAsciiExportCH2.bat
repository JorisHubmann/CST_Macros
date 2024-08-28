'HAsciiExportCH2 Sleeve

Sub Main ()
	Dim Nports,Freq As Integer
	Dim AName, pDimension, Orientation As String

	Nports=2
	Freq=298
	Stepsize=5

	AName="Sleeve"
	pDimension="2D"
	
	Orientation="xy"
	
	Call mASCIIExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)

End Sub

Sub mASCIIExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)

	Field=array("SAR","H")
	
	Dim ExportFolder As String
	ExportFolder= "Z:\CST_Daten\Ergebnisse\DumpData"
	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=" + Freq + ") [CP]") Then
	
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
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_" + Field(1) + "_" + Orientation + ".txt")
			.Mode ("FixedWidth")
			.SetfileType("ascii")
			.UseSubvolume(False)
			.Step(Stepsize)
			.Execute
		End With
	End If

	'Channels
	For i=1 To Nports
		If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=" + Freq + ") [CH_" + i + "]") Then
		
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
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_" + Field(1) + "_" + "CH_" + i + "_" + Orientation + ".txt")
				.Mode ("FixedWidth")
				.SetfileType("ascii")
				.UseSubvolume(False)
				.Step(Stepsize)
				.Execute
			End With
		End If

	Next

End Sub
