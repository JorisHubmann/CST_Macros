Sub Main ()

	Call SExport()

End Sub

Sub SExport()

With SimulationTask
	.Reset
	.Type ("S-Parameters")
	.Name ("Spara")
	.Delete
	.Create
	.Update

	With TOUCHSTONE
   	 	.Reset
    	.FileName ("Z:\CST_Daten\Ergebnisse\DumpData.\Sleeve")
	    .Impedance (50)
	    .FrequencyRange ("Full")
	    .Renormalize (True)
	    .UseARResults (False)
 	   .SetNSamples (0)
 	   .Write
	End With
	
End With


End Sub
