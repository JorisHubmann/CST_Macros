'#Language "WWB-COM"

' H_SAR_AsciiExport
Sub Main ()
	Dim Nports,Freq As Integer

	Dim AName, pDimension, Orientation, ExportFolder As String

	Nports=1
	Stepsize=4

	AName="FD_Val_n2"
	pDimension="3D"
	Orientation=""
	ExportFolder = "Z:\CST_Daten\Simulationen\PaperMitBilgun\Ergebnisse" 'change for directory
	SubV= Array(-250,250,-2,250,-250,250) 'Volume


	Call SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)
	Call HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)

End Sub


Sub SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)

	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\SAR\SAR (f=297.2) [AC1] (10g)") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_SAR" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5))
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If

End Sub

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Sub HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,  ExportFolder,SubV)

	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=297.2) [AC1]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_H" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5))
			'.SetSubvolume(-290,-40,-100,100,-100,200) 'Element Selection (0° und 45°)
			'.SetSubvolume(-290,-40,-100,140,-100,200) 'Element Selection (90°)
			'.SetSubvolume(-250,250,0,250,-250,250) 'PaperBilgun Bilguun simus
			'.SetSubvolume(-270,-20,-250,250,-250,250) 'PaperBilgun my simus
			'.SetSubvolume(-10,10,-100,100,-100,100)
			'.SetSubvolume(-230,-90,-140,140,-110,50) 'PTB Messungen
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If
End Sub

