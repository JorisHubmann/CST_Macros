Option Explicit
'#include "vba_globals_all.lib"
'#include "vba_globals_3d.lib"

' (CSTxMWSxONLY)
' ================================================================================================
' History of Changes
' ------------------
' 28-May-2014 apr: Set excit for sampled case to allow accepted power scaling
' 22-May-2014 apr: fixed Rho matrix issue
' 01-Apr-2014 apr: added SAR for tet mesh
' 22-Jan-2014 twi: r m s  labels removed from dialogue
' 03-Feb-2012 ube: rescaling to user-defined power can be deactivated, UseAR-Flag included (default off)
' 07-Oct-2011 apr: renamed IEEE P1528.1 to IEEE/IEC 62704-1
' 02-Aug-2011 fsr: replaced obsolete 'vba_globals.lib' with 'vba_globals_all.lib' and 'vba_globals_3d.lib'
' 19-Oct-2010 ube: tissue power corrected and new action "Absorbed power"
' 30-Sep-2010 ube: changed label for (0g)-weight to (point)
' 17-Sep-2010 ube: added correct action for subvolumes "sel " + action
' 30-Jul-2010 ube: StoreTemplateSetting included
' 24-Oct-2005 ube: Included into Online Help
' ================================================================================================

' *** global variables
Dim bDSTemplate As Boolean
Public Const sVersion_SARtemplate_CST = "Version 2014-05-28"

Sub DebugTrace(s As String)
'	ReportInformationToWindow(Str$(Now)+" "+s)
End Sub


Public Const sAvgMethod = Array( _
	"IEEE/IEC 62704-1", _
	"IEEE C95.3", _
	"CST C95.3", _
	"CST Legacy", _
	"Constant volume")

Public Const sAction = Array( _
	"Max SAR", _
	"x-Coordinate of Max SAR", _
	"y-Coordinate of Max SAR", _
	"z-Coordinate of Max SAR", _
	"Total SAR", _
	"Max Point SAR", _
	"Tissue volume", _
	"Tissue mass", _
	"Tissue power", _
	"Total absorbed power", _
	"Average power density")

Public Const sVBAaction = Array( _
	"max sar", _
	"max sar x", _
	"max sar y", _
	"max sar z", _
	"total sar", _
	"point sar", _
	"tissue volume", _
	"tissue mass", _
	"tissue power", _
	"power", _
	"average power")


'---------------------------------------------------------------------------------------------------------------------------
Private Function DialogFunction(DlgItem$, Action%, SuppValue&) As Boolean

' -------------------------------------------------------------------------------------------------
' DialogFunction: This function defines the dialog box behaviour. It is automatically called
'                 whenever the user changes some settings in the dialog box, presses any button
'                 or when the dialog box is initialized.
' -------------------------------------------------------------------------------------------------


	If (Action%=1) Or (Action%=2) Then

		If (DlgItem = "Help") Then
			StartHelp "common_preloadedmacro_1D_-_SAR_calculation"
			DialogFunction = True
		End If

		DlgEnable "WgtCustom", IIf((DlgValue("Group_Average") = 3),True,False)
		DlgEnable "extSARname", IIf((DlgValue("CheckExt") = 1),True,False)

		Dim bUserDefPower As Boolean
		If bDSTemplate Then ' this flag is set in Define
			DlgEnable "UseAR", False
			DlgEnable "UserDefinedPower", False
			bUserDefPower = False
		Else
			bUserDefPower = (DlgValue("UserDefinedPower") = 1)
		End If
		DlgEnable "dpower", bUserDefPower
		DlgEnable "AcceptedStimulated", bUserDefPower
		DlgEnable "Text_Wrms", bUserDefPower

		Dim bsubvol As Boolean
		bsubvol = (DlgValue("CheckSubvolume") = 1)
		DlgEnable "xmin", bsubvol
		DlgEnable "ymin", bsubvol
		DlgEnable "zmin", bsubvol
		DlgEnable "xmax", bsubvol
		DlgEnable "ymax", bsubvol
		DlgEnable "zmax", bsubvol
		DlgEnable "Text_dSamplingStep", bsubvol
		DlgEnable "dSamplingStep", bsubvol

		If (DlgItem = "OK") Then
			If bsubvol Then
				If (DlgText("xmin")="" Or DlgText("ymin")="" Or DlgText("zmin")="" Or _
				    DlgText("xmax")="" Or DlgText("ymax")="" Or DlgText("zmax")="" ) Then
					MsgBox "Please check dimensions of subvolume.", vbCritical
					DialogFunction = True   	' There is an error in the settings -> Don't close the dialog box.
				End If
			End If
			If (DlgValue("CheckExt")= 1) Then
				If (DlgText("extSARname")="" ) Then
					MsgBox "Please enter a name extension.", vbCritical
					DialogFunction = True   	' There is an error in the settings -> Don't close the dialog box.
				End If
			End If
		End If

	End If
End Function ' DialogFunction

'---------------------------------------------------------------------------------------------------------------------------
Function Define(sName As String, bCreate As Boolean, bNameChanged As Boolean) As Boolean
DebugTrace(sVersion_SARtemplate_CST + " " + GetApplicationName())
bDSTemplate = Left(GetApplicationName,2)="DS"

Dim iii As Long
Dim nmoni_3dloss As Long
nmoni_3dloss = 0

For iii= 0 To Monitor.GetNumberOfMonitors-1
	If (Monitor.GetMonitorTypeFromIndex(iii) = "Loss density 3D") Then
		nmoni_3dloss = nmoni_3dloss + 1
	End If
Next

If (nmoni_3dloss = 0) Then
	Define = False
	MsgBox "No Power Loss Monitor is defined.", vbCritical
	Exit Function
End If

Dim PwLoss_Names$()
ReDim PwLoss_Names$(nmoni_3dloss-1)

nmoni_3dloss = 0
For iii= 0 To Monitor.GetNumberOfMonitors-1
	If (Monitor.GetMonitorTypeFromIndex(iii) = "Loss density 3D") Then
		PwLoss_Names$(nmoni_3dloss)=Monitor.GetMonitorNameFromIndex(iii)
		nmoni_3dloss = nmoni_3dloss + 1
	End If
Next

Dim aAvgMethod() As String
FillArray aAvgMethod() ,  sAvgMethod

Dim aAction() As String
FillArray aAction() ,  sAction

	Begin Dialog UserDialog 635,378,"SAR Calculation",.DialogFunction ' %GRID:5,3,1,1
		GroupBox 15,6,295,84,"Power Loss Density Monitor and Excitation",.GroupBox4
		DropListBox 40,27,260,192,PwLoss_Names(),.iPwLossmoni
		Text 40,63,90,14,"Excit.String:",.Text1
		TextBox 125,60,170,21,.excit
		GroupBox 15,99,295,117,"Averaging mass",.GroupBox1
		OptionGroup .Group_Average
			OptionButton 40,129,60,15,"10g",.OptionButton1
			OptionButton 120,129,60,15,"1g",.OptionButton2
			OptionButton 195,129,90,15,"Point SAR",.OptionButton3
			OptionButton 40,159,90,15,"Custom:",.OptionButton4
		TextBox 130,157,110,21,.WgtCustom
		Text 250,159,20,15,"g",.Text2
		Text 40, 189, 90, 15, "Accuracy:", .AccText
		TextBox 130, 187,110, 21, .dVolAccuracy
		Text 250,189,20,15,"%",.Text11
		GroupBox 20,219,290,126,"Result value",.GroupBox7
		DropListBox 35,240,260,192,aAction(),.aAction
		OKButton 20,351,90,21
		CancelButton 120,351,90,21
		PushButton 220,351,90,21,"Help",.Help
		GroupBox 320,7,300,49,"SAR result name extension",.GroupBox3
		CheckBox 340,30,25,15,"",.CheckExt
		TextBox 370,27,230,18,.extSARname
		GroupBox 320,63,300,54,"Averaging method",.GroupBox6
		DropListBox 340,90,255,192,aAvgMethod(),.aAvgMethod
		GroupBox 320,126,300,90,"User defined reference power",.GroupBox2
		TextBox 375,144,130,18,.dpower
		Text 515,147,55,15,"W",.Text_Wrms
		OptionGroup .AcceptedStimulated
			OptionButton 340,168,145,15,"accepted (forward)",.OptionButton5
			OptionButton 500,168,95,15,"stimulated",.OptionButton6
		CheckBox 340,144,30,15,"",.UserDefinedPower
		CheckBox 340,192,250,15,"Use AR filter results if available",.UseAR
		GroupBox 320,219,300,126,"Subvolume",.GroupBox5
		CheckBox 340,238,150,15,"Use Subvolume",.CheckSubvolume
		Text 490,238,90,21,"Sampl.Step:",.Text_dSamplingStep
		TextBox 570,236,40,21,.dSamplingStep
		Text 330,258,90,15,"Xmin:",.Text4
		Text 330,300,90,15,"Xmax:",.Text5
		TextBox 330,273,90,21,.xmin
		TextBox 330,315,90,21,.xmax
		Text 425,258,90,15,"Ymin:",.Text6
		Text 425,300,90,15,"Ymax:",.Text7
		TextBox 425,273,90,21,.ymin
		TextBox 425,315,90,21,.ymax
		Text 520,258,90,15,"Zmin:",.Text8
		Text 520,300,90,15,"Zmax:",.Text9
		TextBox 520,273,90,21,.zmin
		TextBox 520,315,90,21,.zmax
		Text 480,360,190,15,sVersion_SARtemplate_CST,.Text10
	End Dialog
	Dim dlg As UserDialog

	' default-settings

	If bDSTemplate Then
		dlg.excit = GetScriptSetting("excit","[AC1]")
	Else
		dlg.excit = GetScriptSetting("excit","[1]")
	End If
	dlg.iPwLossmoni=FindListIndex(PwLoss_Names(), GetScriptSetting("PwMonitor",""))

	dlg.CheckExt = CInt(GetScriptSetting("CheckExt","0"))
	dlg.extSARname = GetScriptSetting("extSARname","")

	dlg.Group_Average = CInt(GetScriptSetting("Group_Average","1"))
	dlg.WgtCustom = GetScriptSetting("WgtCustom","100")
	dlg.dVolAccuracy = GetScriptSetting("dVolAccuracy","0.0001") ' this is in %, so already equal to 62704-1 suggested 0.000001

	dlg.aAvgMethod = FindListIndex(aAvgMethod(), GetScriptSetting("aAvgMethod",aAvgMethod(0)))
	dlg.aAction = FindListIndex(aAction(), GetScriptSetting("aAction",aAction(0)))

	dlg.dpower = GetScriptSetting("dpower","0.5")
	dlg.AcceptedStimulated = CInt(GetScriptSetting("AcceptedStimulated","1"))
	dlg.UseAR = CInt(GetScriptSetting("UseAR","0"))
	dlg.UserDefinedPower = CInt(GetScriptSetting("UserDefinedPower","0"))

	dlg.CheckSubvolume = CInt(GetScriptSetting("CheckSubvolume","0"))
	Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
	Boundary.GetCalculationBox(xmin, xmax, ymin, ymax, zmin, zmax)
	dlg.xmin = GetScriptSetting("xmin",CStr(xmin))
	dlg.ymin = GetScriptSetting("ymin",CStr(ymin))
	dlg.zmin = GetScriptSetting("zmin",CStr(zmin))
	dlg.xmax = GetScriptSetting("xmax",CStr(xmax))
	dlg.ymax = GetScriptSetting("ymax",CStr(ymax))
	dlg.zmax = GetScriptSetting("zmax",CStr(zmax))

	dlg.dSamplingStep = GetScriptSetting("dSamplingStep","0")

	If (dlg.iPwLossmoni<0) Then dlg.iPwLossmoni=0

	If (Not Dialog(dlg)) Then

		' The user left the dialog box without pressing Ok. Assigning False to the function
		' will cause the framework to cancel the creation or modification without storing
		' anything.

		Define = False
	Else

		' The user properly left the dialog box by pressing Ok. Assigning True to the function
		' will cause the framework to complete the creation or modification and store the corresponding
		' settings.

		Define = True

		' Convert the dialog data into strings in order to store them in the script settings database.

		' Determine a proper name for the result item. Changing the name will cause the framework to use
		' the modified name for the result item.

		If (Not bNameChanged) Then
			sName =  "SAR-"+PwLoss_Names(dlg.iPwLossmoni)

		    sName = NoForbiddenFilenameCharacters(sName)
		End If

		' Store the script settings into the database for later reuse by either the define function (for modifications)
		' or the evaluate function.

		StoreScriptSetting("excit",dlg.excit)
		StoreScriptSetting("PwMonitor",PwLoss_Names(dlg.iPwLossmoni))

		StoreScriptSetting("Group_Average",CStr(dlg.Group_Average))
		StoreScriptSetting("WgtCustom",dlg.WgtCustom)
		StoreScriptSetting("dVolAccuracy",dlg.dVolAccuracy)

		StoreScriptSetting("CheckExt",CStr(dlg.CheckExt))
		StoreScriptSetting("extSARname",dlg.extSARname)

		StoreScriptSetting("aAvgMethod",aAvgMethod(dlg.aAvgMethod))
		StoreScriptSetting("aAction",aAction(dlg.aAction))

		StoreScriptSetting("dpower",dlg.dpower)
		StoreScriptSetting("AcceptedStimulated",CStr(dlg.AcceptedStimulated))
		StoreScriptSetting("UseAR",CStr(dlg.UseAR))
		StoreScriptSetting("UserDefinedPower",CStr(dlg.UserDefinedPower))

		StoreScriptSetting("CheckSubvolume",CStr(dlg.CheckSubvolume))
		StoreScriptSetting("xmin",dlg.xmin)
		StoreScriptSetting("ymin",dlg.ymin)
		StoreScriptSetting("zmin",dlg.zmin)
		StoreScriptSetting("xmax",dlg.xmax)
		StoreScriptSetting("ymax",dlg.ymax)
		StoreScriptSetting("zmax",dlg.zmax)
		StoreScriptSetting("dSamplingStep",dlg.dSamplingStep)

		StoreTemplateSetting("TemplateType","0D")

	End If

End Function ' Define

Sub SamplePLD(sLabel As String)
	Dim gscale As Double
	Dim nx As Long, ny As Long, nz As Long
	Dim mass As Double, indexWrite As Long, indexRead As Long, idShape As Long

	Dim sMeshStep As String
	sMeshStep = GetScriptSetting("dSamplingStep","4")
	Dim meshStep As Double, xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double
	meshStep = Val(sMeshStep)
	xmin = GetScriptSetting("xmin","0.0")
	xmax = GetScriptSetting("xmax","0.0")
	ymin = GetScriptSetting("ymin","0.0")
	ymax = GetScriptSetting("ymax","0.0")
	zmin = GetScriptSetting("zmin","0.0")
	zmax = GetScriptSetting("zmax","0.0")
	gscale = Units.GetGeometryUnitToSI()
	DebugTrace("******** Started sampling for SAR, mesh step "+Str$(meshStep*gscale*1000)+" mm ********")

	' generate equidistant hex mesh for desired volume
	nx = (xmax-xmin)/meshStep +1
	ny = (ymax-ymin)/meshStep +1
	nz = (zmax-zmin)/meshStep +1
	Dim xpos() As Double, ypos() As Double, zpos() As Double
	ReDim xpos(nx)
	ReDim ypos(ny)
	ReDim zpos(nz)

	Dim i As Long
	For i = 0 To nx-1
		xpos(i) = xmin*(nx-i-1)/(nx-1) + xmax*i/(nx-1)
	Next
	For i = 0 To ny-1
		ypos(i) = ymin*(ny-i-1)/(ny-1) + ymax*i/(ny-1)
	Next
	For i = 0 To nz-1
		zpos(i) = zmin*(nz-i-1)/(nz-1) + zmax*i/(nz-1)
	Next

	' write mesh to standard FIT hex mesh file
	ResultOperator3D.initialize(nx,ny,nz,"hexmesh")
	For i = 0 To nx-1
		ResultOperator3D.setxre(i, xpos(i)*gscale)
	Next
	For i = 0 To ny-1
		ResultOperator3D.setyre(i, ypos(i)*gscale)
	Next
	For i = 0 To nz-1
		ResultOperator3D.setzre(i, zpos(i)*gscale)
	Next
	ResultOperator3D.Save(GetProjectPath("Result")+sLabel+"_HexMesh.msh")
	With Resulttree
		.Name sLabel+"_HexMesh.msh"
		.File GetProjectPath("Result")+sLabel+"_HexMesh.msh"
		.YLabel sLabel+"_HexMesh.msh"
		.Type "hidden"
		.Add
	End With

	' Get power loss density at grid cell centers from currently selected result in tree
	DebugTrace("******** Interpolate power loss density to hex grid with "+Str$(nx)+" *"+Str$(ny)+" *"+Str$(nz)+" ="+Str$(nx*ny*nz)+" nodes... ********")
	Dim xarr() As Double, yarr() As Double, zarr() As Double
	ReDim xarr((nx-1)*(ny-1)*(nz-1)-1)
	ReDim yarr((nx-1)*(ny-1)*(nz-1)-1)
	ReDim zarr((nx-1)*(ny-1)*(nz-1)-1)
	Dim ix As Integer, iy As Integer, iz As Integer
	For iz = 0 To nz-2
	For iy = 0 To ny-2
	For ix = 0 To nx-2
		indexWrite = ix+iy*(nx-1)+iz*(nx-1)*(ny-1)
		xarr(indexWrite) = 0.5*(xpos(ix)+xpos(ix+1))
		yarr(indexWrite) = 0.5*(ypos(iy)+ypos(iy+1))
		zarr(indexWrite) = 0.5*(zpos(iz)+zpos(iz+1))
	Next ix
	Next iy
	Next iz
	VectorPlot3D.SetPoints(xarr, yarr, zarr)
	VectorPlot3D.CalculateList

	DebugTrace("******************** Build mass density matrix ********************")
	Solid.GenerateShapeIdTable
	Solid.SetMissingAsWarning True
	Dim nShapeIDs As Long
	nShapeIDs = Solid.GetNumberOfShapeIds
	Dim rho() As Double
	Dim dRho As Double
	Dim strName As String, matname As String
	ReDim rho(nShapeIDs)
	For i = 0 To nShapeIDs-1
		strName = Solid.GetShapeNameFromID(i)
		strName = Mid(strName, 12)
		matname = ""
		If InStr(strName, "air_0:brick")=0 Then
			matname = Solid.GetMaterialNameForShape(strName)
		End If
		dRho = 0.0
		If matname<>"" Then
			Material.GetRho(matname, dRho)
		End If
		rho(i) = dRho
	Next i
	DebugTrace(Str$(nShapeIDs)+" ShapeIDs found, get shape ids for sample positions")
	VectorPlot3D.determineShapeIDs

	Dim loss() As Double
	loss = VectorPlot3D.GetList("xre")
	Dim shapeArray() As Long
	shapeArray = VectorPlot3D.GetShapeList()

	' write power loss density in standard FIT hex result file and mass density into standard FIT hex matrix file
	mass = 0
	DebugTrace("***************** Write loss density and mass density *****************")
	Dim loss3d_array() As Double, rho3d_array() As Double
	ReDim loss3d_array(nx*ny*nz-1)  ' -1 to avoid failure because of wrong dimension in SetArray
	ReDim rho3d_array(nx*ny*nz-1)
	For iz = 0 To nz-1
	For iy = 0 To ny-1
	For ix = 0 To nx-1
		indexWrite = ix+iy*nx+iz*nx*ny
		If ix<nx-1 And iy<ny-1 And iz<nz-1 Then
			indexRead = ix+iy*(nx-1)+iz*(nx-1)*(ny-1)
			loss3d_array(indexWrite) = loss(indexRead)
			idShape = shapeArray(indexRead)
			If idShape<nShapeIDs+1 Then
				rho3d_array(indexWrite) = rho(idShape)
				mass = mass + rho(idShape)
			Else
				Debug.Print indexRead; idShape; indexWrite
				rho3d_array(indexWrite) = 0
			End If
		Else
			loss3d_array(indexWrite) = 0
			rho3d_array(indexWrite) = 0
		End If
	Next ix
	Next iy
	Next iz
	DebugTrace("total mass "+Str$(mass*(xpos(nx-1)-xpos(0))*(ypos(ny-1)-ypos(0))*(zpos(nz-1)-zpos(0))*gscale*gscale*gscale/(nx-1)/(ny-1)/(nz-1))+" kg")
	Dim loss3d As Object, rho3d As Object
	Set loss3d = Result3D("")
	loss3d.initialize(nx,ny,nz,"scalar")
	loss3d.settype("dynamic loss density")
	loss3d.SetArray(loss3d_array, "xre")
	Dim frequency As Double
	frequency = GetFieldFrequency()
	DebugTrace("Frequency "+Str$(frequency))
	loss3d.SetFrequency(frequency)
	loss3d.save(GetProjectPath("Result")+sLabel+"_PLD.m3d")
	With Resulttree
		.Name sLabel+"_PLD.m3d"
		.File GetProjectPath("Result")+sLabel+"_PLD.m3d"
		.YLabel sLabel+"_PLD.m3d"
		.Type "hidden"
		.Add
	End With
	ResultOperator3D.initialize(nx,ny,nz,"vector")
	ResultOperator3D.SetArray(rho3d_array, "xre")
	ResultOperator3D.save(GetProjectPath("Result")+sLabel+"_Rho.mat")
	With Resulttree
		.Name sLabel+"_Rho.mat"
		.File GetProjectPath("Result")+sLabel+"_Rho.mat"
		.YLabel sLabel+"_Rho.mat"
		.Type "hidden"
		.Add
	End With
End Sub ' SamplePLD

Function SettingsChanged(sarname As String) As Boolean

	SettingsChanged = True

	Dim sSettingsFile As String
	sSettingsFile = GetProjectPath("Result")+sarname+"_settings.txt"
	If Dir(sSettingsFile)="" Then
		Exit Function
	End If

	Dim sLine As String, sName As String
	Open sSettingsFile For Input As #1
	While Not EOF(1)
		Line Input #1, sLine
		sName = Left$(sLine, InStr(sLine, ", ")-1)
		sLine = Mid$(sLine, InStr(sLine, " ")+1)
		If GetScriptSetting(sName, "")<>sLine Then
			DebugTrace(sName+" has changed from "+sLine+" to "+GetScriptSetting(sName, ""))
			Close #1
			Exit Function
		End If
	Wend
	Close #1

	DebugTrace("All settings unchanged.")
	SettingsChanged = False
End Function ' SettingsChanged

Sub SaveSettings(sarname As String)
	Dim sSettingsFile As String
	sSettingsFile = GetProjectPath("Result")+sarname+"_settings.txt"
	DebugTrace("Write settings to "+sSettingsFile)
	Open sSettingsFile For Output As #1
	Print #1, "xmin, "; GetScriptSetting("xmin","")
	Print #1, "ymin, "; GetScriptSetting("ymin","")
	Print #1, "zmin, "; GetScriptSetting("zmin","")
	Print #1, "xmax, "; GetScriptSetting("xmax","")
	Print #1, "ymax, "; GetScriptSetting("ymax","")
	Print #1, "zmax, "; GetScriptSetting("zmax","")
	Print #1, "dSamplingStep, "; GetScriptSetting("dSamplingStep","0")
	Close #1
	With Resulttree
		.Name sarname+"_settings.txt"
		.File GetProjectPath("Result")+sarname+"_settings.txt"
		.YLabel sarname+"_settings.txt"
		.Type "hidden"
		.Add
	End With
End Sub ' SaveSettings

Function Evaluate0D() As Double
	ReportInformationToWindow("Calculate SAR Result: "+GetScriptSetting("aAction","")+" ...")
	DebugTrace("SAR Result, "+sVersion_SARtemplate_CST + " " + GetApplicationName())
	bDSTemplate = Left(GetApplicationName,2)="DS"

	Dim sExcit As String
	If bDSTemplate Then
		sExcit = " " + GetScriptSetting("excit","[AC1]")
	Else
		sExcit = " " + GetScriptSetting("excit","[1]")
	End If

	Dim pwname As String
	pwname = GetScriptSetting("PwMonitor","")

	Mesh.ViewMeshMode  False

	With SAR
		.Reset
		.PowerlossMonitor pwname + sExcit

		Dim sWeight As String
		Select Case CInt(GetScriptSetting("Group_Average","0"))
		Case 0
				sWeight="10.0"
		Case 1
				sWeight="1.0"
		Case 2
				sWeight="0.0"
		Case 3
				sWeight=GetScriptSetting("WgtCustom","100")
		End Select

		.AverageWeight Evaluate(sWeight)

		If (CInt(GetScriptSetting("CheckSubvolume","0"))=0) Then
			' do not use any subvolume specification
			.SetOption "no subvolume"
		Else
			' limits the SAR calculation on the user defined volume to save execution time.
			.SetOption "subvolume only"
			.Volume Evaluate(GetScriptSetting("xmin","0.0")), Evaluate(GetScriptSetting("xmax","0.0")), Evaluate(GetScriptSetting("ymin","0.0")), Evaluate(GetScriptSetting("ymax","0.0")), Evaluate(GetScriptSetting("zmin","0.0")), Evaluate(GetScriptSetting("zmax","0.0"))
		End If

		.SetOption "volaccuracy "+CStr(0.01*Evaluate(GetScriptSetting("dVolAccuracy","0.0001"))) ' this is in %, so already equal to 62704-1 suggested 0.000001
		.SetOption GetScriptSetting("aAvgMethod","")

		If (CInt(GetScriptSetting("UserDefinedPower","0"))=0) Then
			.SetOption "no rescale"
		Else
			.SetOption "rescale "+Evaluate(GetScriptSetting("dpower","0.5"))
			'
			If (CInt(GetScriptSetting("AcceptedStimulated","1"))=0) Then
				.SetOption "scale accepted"
			Else
				.SetOption "scale stimulated"
			End If
		End If

		If (CInt(GetScriptSetting("UseAR","1"))=0) Then
			.SetOption "no ar results"
		Else
			.SetOption "use ar results"
		End If

		Dim sarname As String, stmp As String, sLabelSAR As String

		stmp = CStr(Evaluate(sWeight))

		If stmp <> "0" Then
			sLabelSAR = "(" + stmp + "g)"
		Else
			sLabelSAR = "(Point)"
		End If

		sarname = Replace(pwname,"loss ","SAR ") + sExcit + " " + sLabelSAR

		If (CInt(GetScriptSetting("CheckExt","0"))=1) Then
			sarname = sarname +	GetScriptSetting("extSARname","")
		End If

		If GetScriptSetting("dSamplingStep","0")<>0 Then
			ReportInformationToWindow("Note: SAR calculation with sampled power loss density is a preview (beta) functionality. Results can not be displayed in 2D/3D.")
			If Dir(GetProjectPath("Result")+sarname+"_HexMesh.msh")="" Or _
				Dir(GetProjectPath("Result")+sarname+"_PLD.m3d")="" Or _
				SettingsChanged(sarname) _
			Then
				Dim sTreePathPLD As String
				sTreePathPLD = "2D/3D Results\Power Loss Dens.\"+pwname+sExcit
				If Not SelectTreeItem(sTreePathPLD) Then
					Dim sErrIDNE As String ' avoid double output of the error msg
					sErrIDNE = "Tree item does not exist: '"+sTreePathPLD+"'"
					ReportError(sErrIDNE)
				End If
				SamplePLD(sarname)
				SaveSettings(sarname) ' Save settings after successful completion
			End If
			SAR.PowerlossMonitor(GetProjectPath("Result")+sarname+"_PLD.m3d")
			SAR.SetOption("MeshFileName "+GetProjectPath("Result")+sarname+"_HexMesh") '.msh is assumed automatically
			SAR.SetOption("RhoFileName "+sarname+"_Rho") 'result path and .mat is assumed automatically
			Dim sExcitRaw As String
			sExcitRaw = Mid$(sExcit, 3, Len(sExcit)-3)
			SAR.SetOption("Excitation "+sExcitRaw)
		End If
		.SetLabel sarname
		.SetOption ("opt sarcubes")		' ABI: Output SAR cubes activated

		' apr: SAR calculation now automatically checkes whether a result has already been calculated with the given settings
		.Create

		If GetScriptSetting("dSamplingStep","0")<>0 Then
			If Dir(GetProjectPath("Result")+sarname+".m3d")<>"" Then
				' 3D result has to be renamed because it causes trouble in visualization
				If Dir(GetProjectPath("Result")+sarname+"_SAR.m3d")<>"" Then
					Kill GetProjectPath("Result")+sarname+"_SAR.m3d"
				End If
'				DebugTrace("Rename '"+GetProjectPath("Result")+sarname+".m3d"+"' to '...\"+sarname+"_SAR.m3d'")
'				Name GetProjectPath("Result")+sarname+".m3d" As GetProjectPath("Result")+sarname+"_SAR.m3d" ' there seems to be a problem with linux
				DebugTrace("Copy '"+GetProjectPath("Result")+sarname+".m3d"+"' to '...\"+sarname+"_SAR.m3d'")
				FileCopy GetProjectPath("Result")+sarname+".m3d", GetProjectPath("Result")+sarname+"_SAR.m3d"
				DebugTrace("Kill '"+GetProjectPath("Result")+sarname+".m3d'")
				Kill GetProjectPath("Result")+sarname+".m3d"
			End If
		End If

		Dim aAction() As String
		FillArray aAction() ,  sAction

		Dim aVBAaction() As String
		FillArray aVBAaction() ,  sVBAaction

		Dim iaction As Integer
		iaction = FindListIndex(aAction(), GetScriptSetting("aAction",aAction(0)))

		If (CInt(GetScriptSetting("CheckSubvolume","0"))=0) Then
			' do not use any subvolume specification
			Evaluate0D = .GetValue(aVBAaction(iaction))
		Else
			Evaluate0D = .GetValue("sel "+aVBAaction(iaction))
		End If

	End With
End Function ' Evaluate0D

Function Evaluate1D() As Object
	Set Evaluate1D = Result1D("")
	With Evaluate1D
		.Initialize 1
		.SetXY 0, 1, Evaluate0D()
	End With
End Function

Sub Main0d
	ActivateScriptSettings True
	ClearScriptSettings
	If (Define("test", True, False)) Then
		MsgBox CStr(Evaluate0D())
	End If
	ActivateScriptSettings False
End Sub

Sub Main1d
	ActivateScriptSettings True
	ClearScriptSettings
	If (Define("test", True, False)) Then
		Dim stmpfile As String
		stmpfile = "Test1D_tmp.txt"
		Dim r1d As Object
		Set r1d = Evaluate1D
		r1d.Save stmpfile
		r1d.AddToTree "1D Results\Test 1D"
		SelectTreeItem "1D Results\Test 1D"
		With Resulttree
		    .UpdateTree
		    .RefreshView
		End With
	End If
	ActivateScriptSettings False
End Sub
