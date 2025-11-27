'#Language "WWB-COM"

' H_SAR_AsciiExport
Sub Main ()
	Dim Nports,Freq As Integer

	Dim AName, pDimension, Orientation, ExportFolder As String

	Nports=1
	Stepsize=4

	AName="IM_20mm"
	pDimension="3D"
	Orientation=""
	ExportFolder = "Z:\CST_Daten\Simulationen\PaperMitBilgun\Ergebnisse" 'change for directory
	SubV= Array(-272,-20,-250,250,-250,250) 'Volume

	Call SARExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)
	Call HExport(Nports,Freq,Stepsize, AName, pDimension, Orientation, ExportFolder,SubV)

End Sub

Sub CalcSAR(Nports,Freq)
	Dim i As Double
	Dim objN As String
	Dim averagingMethod As String

	averagingMethod="IEEE/IEC 62704-1"

	'CP-Mode
	objN = "SAR_CP"
	If Resulttree.DoesTreeItemExist("2D/3D Results\SAR\" & objN ) Then
		With Resulttree
			.Name ("2D/3D Results\SAR\" & objN )
			.Delete
		End With
	End If
	If Not Resulttree.DoesTreeItemExist("2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor("loss (f=297.2) [AC1]")
			.AverageWeight(10)
			.SetLabel("SAR (f=297.2) [SAR_AC1] (10g)")
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
		End With

	End If
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

Sub PLDExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,  ExportFolder,SubV)

	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\Power Loss Dens.\loss (f=297.2) [AC1]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_PLD" + Orientation + ".h5")
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

Sub EExport(Nports,Freq,Stepsize, AName, pDimension, Orientation,  ExportFolder,SubV)

	'MkDir ExportFolder
	'CP-Mode
	If SelectTreeItem ("2D/3D Results\E-Field\e-field (f=297.2) [AC1]") Then

		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (ExportFolder + ".\" + AName + "_" + pDimension + "_E" + Orientation + ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(SubV(0),SubV(1),SubV(2),SubV(3),SubV(4),SubV(5))
			'SetSubvolume(-290,-40,-100,100,-100,200) 'Element Selection
			'.SetSubvolume(-10,10,-100,100,-100,100)
			.UseSubvolume(True)
			.Step(Stepsize)
			.Execute
		End With
	End If
End Sub
