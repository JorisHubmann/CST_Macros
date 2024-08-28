' H_SAR_AsciiExport 2 Sleeve

Sub Main ()
	Dim Nports,Freq As Integer
	Dim AName, pDimension, Orientation	As String

	Nports=16
	Freq=298
	Stepsize=1

	AName="Sleeve" 'change to the correct name
	pDimension="2D" 
	Orientation="xy" 'change to the correct orientation
	
	Call SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)
	Call HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)
End Sub


'no changes to this part needed:
Sub SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)

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
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR_" + Orientation + ".txt")
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
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR_" + "CH_" + i + "_" + Orientation + ".txt")
				.Mode ("FixedWidth")
				.SetfileType("ascii")
				.UseSubvolume(False)
				.Step(Stepsize)
				.Execute
			End With
		End If

	Next

End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation)
	
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
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H_" + Orientation + ".txt")
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
				.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H_" + "CH_" + i + "_" + Orientation + ".txt")
				.Mode ("FixedWidth")
				.SetfileType("ascii")
				.UseSubvolume(False)
				.Step(Stepsize)
				.Execute
			End With
		End If

	Next

End Sub