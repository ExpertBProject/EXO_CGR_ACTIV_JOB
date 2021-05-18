Module Job

#Region "Método Principal"

    Public Sub Main()

        Dim log As EXO_Log.EXO_Log
        log = New EXO_Log.EXO_Log("C:\Temp\LogActividades\Log.txt", 50)

        Dim iCountExeJob As Integer = 0


        If iCountExeJob = 0 Then
            Procesos.EnviarActividad(log)
        End If

    End Sub

#End Region

End Module
