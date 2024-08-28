'#Language "WWB-COM"

Option Explicit

Sub Main

Dim i As Double

Dim Nports, dp As Integer

Dim Phase As Variant

Dim Amp As Variant

Phase = Array(-44.09103,-31.33609,-254.7872,-288.4487,-349.5324,-262.9757,-123.9512,-250.1151,-261.3318,-241.4549,-261.9591,-336.8515,-291.9968,-152.4167,-334.8747,-246.5339)

Amp = Array(0.48679,0.43586,0.44678,0.30635,0.50851,0.51077,0.81763,0.79483,0.64432,0.37861,0.81158,0.53283,0.35073,0.939,0.87594,0.55016)

Nports = 16

	For i=1 To Nports

		With SimulationTask

			.reset

			.name("Rnd_7")

			.type("AC")

			.SetComplexPortExcitation(i, Amp(i-1), Phase(i-1))

			.SetPortSourceType(i, "Signal")

		End With

	Next

End Sub
