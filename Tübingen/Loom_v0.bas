'#Language "WWB-COM"

Option Explicit

Sub Main

	Call Solver

End Sub

Sub Solver


	Dim Ch, Cm, Cl, te As Double
	Dim w, l, g As Single
    Dim ftread As String, fl As String
    Dim param As Variant
    Dim S, i As Integer
    Dim goalID As Long

	S=1
	i=1

	While S = 1

    	ftread = "D:\anikulin\CST\B1_Loom\Fo8\Loose\Scale\param.txt"
    	Open ftread For Input As #1
    	fl = Input$(LOF(1), 1)
    	Close #1
    	param = Split(fl, vbNewLine)

		'MsgBox param(1)

		Ch = param(0)
		StoreParameter("Ch", Ch)

		Cm = param(1)
		StoreParameter("Cm", Cm)

		Cl = param(2)
		StoreParameter("Cl", Cl)

		w = param(3)
		StoreParameter("w", w)

		l = param(4)
		StoreParameter("l", l)

		g = param(5)
		StoreParameter("g", g)

		S = param(6)

		RebuildOnParametricChange(True, False)

		SolverParameter.SolverType "Frequency domain solver", "Tetrahedral"

		FDSolver.Start

		With Optimizer

			.Start

		End With


			If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=400) [1]") Then
			    Plot3DPlotsOn2DPlane True
			    Wait(1)
				ScalarPlot2D.PlaneNormal ("z")
				ScalarPlot2D.PlaneCoordinate (0)
				With ASCIIExport

					.Reset
					.FileName ("D:\anikulin\CST\B1_Loom\Fo8\Loose\Scale\H_cfg1_" & i & ".txt")
					.Mode("FixedNumber")
					.SetfileType("ascii")
					.UseSubvolume(True)
					.SetSubvolume(-85,85,-85,85,0,0)
					.Step(200)
					.Execute

				End With

			End If

		te=Ch

		Ch=Cl

		Cl=Ch

		RebuildOnParametricChange(True, False)

		SolverParameter.SolverType "Frequency domain solver", "Tetrahedral"

		FDSolver.Start

			If SelectTreeItem ("2D/3D Results\H-Field\h-field (f=400) [1]") Then
			    Plot3DPlotsOn2DPlane True
			    Wait(1)
				ScalarPlot2D.PlaneNormal ("z")
				ScalarPlot2D.PlaneCoordinate (0)
				With ASCIIExport

					.Reset
					.FileName ("D:\anikulin\CST\B1_Loom\Fo8\Loose\Scale\H_cfg2_" & i & ".txt")
					.Mode("FixedNumber")
					.SetfileType("ascii")
					.UseSubvolume(True)
					.SetSubvolume(-85,85,-85,85,0,0)
					.Step(200)
					.Execute

				End With

			End If

		Save

		i=i+1

		'Reset

	Wend


End Sub
