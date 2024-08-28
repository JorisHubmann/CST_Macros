'#Language "WWB-COM"

Option Explicit

Sub Main

	Dim i As Double

	Dim Nports As Integer

	Nports = 7

		For i=1 To Nports

			'Port.NewFolder i
			'LumpedElement.NewFolder i

			'With Port

				'Port.RenameItem "port" & i, i &":port" & i
				'Port.RenameItem "port" & i+10, i &":port" & i+10

			'End With

			With Port

				LumpedElement.Rename "10:element" & i+7 &"_4", "10:" & i
				'LumpedElement.Rename "port" & i+10, i &":port" & i+10

			End With

		Next

End Sub
