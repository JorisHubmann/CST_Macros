'#Language "WWB-COM"

Option Explicit
'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

Sub Main
	Dim Nports,Freq As Integer
	Nports = 16
	Freq = 300
	Dim xmin, xmax, ymin, ymax, zmin, zmax As Integer
	Dim stepsize As Integer

	xmin = -120
	xmax =  120
	ymin = -135
	ymax =  135
	zmin = -165
	zmax =  105
    stepsize = 2

	Call createModes(Nports,Freq)
	Call calcSAR(Nports,Freq,xmin,xmax,ymin,ymax,zmin,zmax)
	Call exportSAR(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
	Call exportCPSAR(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
	Call exportHfield(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
	Call exportEfield(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
End Sub


Sub createModes(Nports,Freq)
Dim i, j As Double

'# Individual Channels
For i=1 To Nports
	With SimulationTask
		.reset
		.name("CH" & i)
		.type("AC")
		If Not .DoesExist Then
			.create
			.ResetCombineMonitorFilters
			.AddCombineMonitorFilter("loss (f=" & Freq & ")")
			.SetProperty("docombineresults",1)
			.SetProperty("blocknameforcombineresults","MWSSCHEM1")
			.SetComplexPortExcitation(i, 1, 0)
			.SetPortSourceType(i, "Signal")
			.Update
		End If
		.SetProperty ("enabled", "False")
	End With
Next

For i=1 To Nports
	For j=i+1 To Nports
		With SimulationTask
'# Mode 1
			.reset
			.name("Ch" & i & "-" & j & "_000deg")
			.type("AC")
			If Not .DoesExist Then
				.create
				.ResetCombineMonitorFilters
				.AddCombineMonitorFilter("loss (f=" & Freq & ")")
				.SetProperty("docombineresults",1)
				.SetProperty("blocknameforcombineresults","MWSSCHEM1")
				.SetComplexPortExcitation(i, 1, 0)
				.SetComplexPortExcitation(j, 1, 0)
				.SetPortSourceType(i, "Signal")
				.SetPortSourceType(j, "Signal")
				.Update
			End If
			.SetProperty ("enabled", "False")
'# Mode 2
			.reset
			.name("Ch" & i & "-" & j & "_090deg")
			.type("AC")
			If Not .DoesExist Then
				.create
				.ResetCombineMonitorFilters
				.AddCombineMonitorFilter("loss (f=" & Freq & ")")
				.SetProperty("docombineresults",1)
				.SetProperty("blocknameforcombineresults","MWSSCHEM1")
				.SetComplexPortExcitation(i, 1, 0)
				.SetComplexPortExcitation(j, 1, 90)
				.SetPortSourceType(i, "Signal")
				.SetPortSourceType(j, "Signal")
				.Update
			End If
			.SetProperty ("enabled", "False")
		End With
	Next
Next

' CP-Mode
With SimulationTask
		.reset
		.name("CP_Mode")
		.type("AC")
		If Not .DoesExist Then
			.create
			.ResetCombineMonitorFilters
			.AddCombineMonitorFilter("loss (f=" & Freq & ")")
			.SetProperty("docombineresults",1)
			.SetProperty("blocknameforcombineresults","MWSSCHEM1")
			For i =1 To Nports
				.SetComplexPortExcitation(i, 1, (i-1) * (-360/(Nports)))
				.SetPortSourceType(i, "Signal")
			Next
			.Update

		End If
		.SetProperty ("enabled", "False")
	End With
End Sub

Sub calcSAR(Nports,Freq,xmin,xmax,ymin,ymax,zmin,zmax)
Dim i, j As Double
'Dim Nports,Freq As Integer

Dim objN As String
Dim averagingMethod As String
' averagingMethod = "CST 62704-1"
' averagingMethod = "IEEE C95.3"
averagingMethod = "CST C95.3"
' averagingMethod = "CST Legacy"
'Nports = 2
'Freq = 300
'# Individual Channels
For i=1 To Nports
	objN = "SAR_Ch" & i
	If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor("loss (f=" & Freq & ") [CH" & i & "]")
			.AverageWeight(10)
			.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
			.SetOption("subvolume only")
			.SetLabel(objN)
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
		End With
	End If
Next

For i=1 To Nports
	For j=i+1 To Nports
		'# Mode 1
		objN = "SAR_Ch" & i & "-" & j & "_000deg"
		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor("loss (f=" & Freq & ") [Ch" & i & "-" & j & "_000deg]")
			.AverageWeight(10)
			.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
			.SetOption("subvolume only")
			.SetLabel(objN)
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			'#Err.Raise vbObjectError + 513, "MyProj.MyObject", "loss (f=" & Freq & ") [Ch" & i & "-" & j & "_000deg]"
			'#Err.Raise("loss (f=" & Freq & ") [Ch" & i & "-" & j & "_000deg")
			.Create
		End With
		End If
		'# Mode 2
		objN = "SAR_Ch" & i & "-" & j & "_090deg"
		If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
		With SAR
			.reset
			.PowerlossMonitor("loss (f=" & Freq & ") [Ch" & i & "-" & j & "_090deg]")
			.AverageWeight(10)
			.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
			.SetOption("subvolume only")
			.SetLabel(objN)
			'# set to 1W stimulated power
			.SetOption("scale stimulated")
			.SetOption("rescale 1")
			.SetOption (averagingMethod)
			.Create
		End With
		End If
	Next
Next

'CP-Mode
objN = "SAR_CP_Mode"
If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
	With SAR
		.reset
		.PowerlossMonitor("loss (f=" & Freq & ") [CP_Mode]")
		.AverageWeight(10)
		.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
		.SetOption("subvolume only")
		.SetLabel("SAR_CP_Mode")
		'# set to 1W stimulated power
		.SetOption("scale stimulated")
		.SetOption("rescale 1")
		.SetOption (averagingMethod)
		.Create
	End With
End If
End Sub


Sub exportSAR(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
Dim i, j As Double
Dim sExportFolder As String
'Dim stepsize As Double
'stepsize = 2
sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d_SAR_" & stepsize & "mm\"
CST_MkDir sExportFolder

'# Individual Channels
For i=1 To Nports
	If SelectTreeItem ("2D/3D Results\SAR\SAR_Ch" & i) Then
		Plot3DPlotsOn2DPlane False
		Wait(1)
		With ASCIIExport
			.Reset
			.FileName (sExportFolder + ".\SAR_Ch" &i & ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
			.UseSubvolume(True)
			.Step(stepsize)
			.Execute
		End With
	End If

Next

For i=1 To Nports
	For j=i+1 To Nports
		' Mode 1
		If SelectTreeItem ("2D/3D Results\SAR\SAR_Ch" & i & "-" & j & "_000deg") Then
	    	Plot3DPlotsOn2DPlane False
	    	Wait(1)
			With ASCIIExport
				.Reset
				.FileName (sExportFolder + ".\SAR_Ch" & i & "-" & j & "_000deg.h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
				.UseSubvolume(True)
				.Step(stepsize)
				.Execute
			End With
		End If
		' Mode 2
		If SelectTreeItem ("2D/3D Results\SAR\SAR_Ch" & i & "-" & j & "_090deg") Then
		    Plot3DPlotsOn2DPlane False
		    Wait(1)
			With ASCIIExport
				.Reset
				.FileName (sExportFolder + ".\SAR_Ch" & i & "-" & j & "_090deg.h5")
				.Mode ("FixedWidth")
				.SetfileType("hdf5")
				.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
				.UseSubvolume(True)
				.Step(stepsize)
				.Execute
			End With
		End If
	Next
Next

'CP-Mode
If SelectTreeItem ("2D/3D Results\SAR\SAR_CP_Mode") Then
	Plot3DPlotsOn2DPlane False
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (sExportFolder + ".\SAR_CP_Mode.h5")
		.Mode ("FixedWidth")
		.SetfileType("hdf5")
		.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
		.UseSubvolume(True)
		.Step(stepsize)
		.Execute
	End With
End If

End Sub

Sub exportSAR2(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax)
Dim i, j As Double
Dim sExportFolder As String
Dim stepsize As Double
stepsize = 2
sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d_SAR\"
CST_MkDir sExportFolder

'CP-Mode
If SelectTreeItem ("2D/3D Results\SAR\SAR_CP_Mode") Then
	Plot3DPlotsOn2DPlane False
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (sExportFolder + ".\SAR_CP_Mode.h5")
		.Mode ("FixedWidth")
		.SetfileType("hdf5")
		.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
		.UseSubvolume(True)
		.Step(stepsize)
		.Execute
	End With
End If

End Sub


Sub exportEfield(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
Dim i, j As Double
Dim sExportFolder As String
'Dim stepsize As Double
'stepsize = 2
sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d_SAR_" & stepsize & "mm\"
CST_MkDir sExportFolder

'# Individual Channels
For i=1 To Nports
	If SelectTreeItem ("2D/3D Results\E-Field\e-field (f=" & Freq & ") [CH_" & i & "]") Then
	    Plot3DPlotsOn2DPlane False
	    Wait(1)
		With ASCIIExport
			.Reset
			.FileName (sExportFolder + ".\E_Ch" &i & ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
			.UseSubvolume(True)
			.Step(stepsize)
			.Execute
		End With
	End If
Next

End Sub

Sub exportHfield(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
Dim i, j As Double
Dim sExportFolder As String
'Dim stepsize As Double
'stepsize = 2
sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d_SAR_" & stepsize & "mm\"
CST_MkDir sExportFolder

'# Individual Channels
For i=1 To Nports
	If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=" & Freq & ") [CH_" & i & "]") Then
	    Plot3DPlotsOn2DPlane False
	    Wait(1)
		With ASCIIExport
			.Reset
			.FileName (sExportFolder + ".\H_Ch" &i & ".h5")
			.Mode ("FixedWidth")
			.SetfileType("hdf5")
			.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
			.UseSubvolume(True)
			.Step(stepsize)
			.Execute
		End With
	End If
Next

End Sub

Sub exportCPSAR(Nports, Freq,xmin,xmax,ymin,ymax,zmin,zmax,stepsize)
Dim i, j As Double
Dim sExportFolder As String
'Dim stepsize As Double
'stepsize = 2
sExportFolder = GetProjectPathMaster_LIB() + "\Export\3d_SAR_" & stepsize & "mm\"
CST_MkDir sExportFolder


' Create mode
With SimulationTask
		.reset
		.name("CP_Mode")
		.type("AC")
		If Not .DoesExist Then
			.create
			.ResetCombineMonitorFilters
			.AddCombineMonitorFilter("loss (f=" & Freq & ")")
			.SetProperty("docombineresults",1)
			.SetProperty("blocknameforcombineresults","MWSSCHEM1")
			For i =1 To Nports
				.SetComplexPortExcitation(i, 1, (i-1) * (-360/(Nports)))
				.SetPortSourceType(i, "Signal")
			Next
			.Update

		End If
		.SetProperty ("enabled", "False")
	End With
' calc SAR
Dim objN As String
objN = "SAR_CP_Mode"
Dim averagingMethod As String
averagingMethod = "CST C95.3"

If Not Resulttree.DoesTreeItemExist( "2D/3D Results\SAR\" & objN ) Then
	With SAR
		.reset
		.PowerlossMonitor("loss (f=" & Freq & ") [CP_Mode]")
		.AverageWeight(10)
		.Volume(xmin,xmax,ymin,ymax,zmin,zmax)
		.SetOption("subvolume only")
		.SetLabel("SAR_CP_Mode")
		'# set to 1W stimulated power
		.SetOption("scale stimulated")
		.SetOption("rescale 1")
		.SetOption (averagingMethod)
		.Create
	End With
End If

'export
If SelectTreeItem ("2D/3D Results\SAR\SAR_CP_Mode") Then
	Plot3DPlotsOn2DPlane False
	Wait(1)
	With ASCIIExport
		.Reset
		.FileName (sExportFolder + ".\SAR_CP_Mode.h5")
		.Mode ("FixedWidth")
		.SetfileType("hdf5")
		.SetSubvolume(xmin,xmax,ymin,ymax,zmin,zmax)
		.UseSubvolume(True)
		.Step(stepsize)
		.Execute
	End With
End If

End Sub
