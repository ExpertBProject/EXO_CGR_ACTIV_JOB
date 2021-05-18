Module Job

#Region "Método Principal"

    Public Sub Main()

        Dim log As EXO_Log.EXO_Log
        log = New EXO_Log.EXO_Log("C:\Temp\LogActividades\Log.txt", 50)

        Dim iCountExeJob As Integer = 0

        'Comprobamos que el JOB no está en ejecución y lo lanzamos
        '    For Each oProcess As Process In Process.GetProcesses()
        '         If Left(oProcess.ProcessName.ToString, 12) = "EXO_CGR_ACTIV_JOB" Then
        '              iCountExeJob += 1
        '           End If
        '        Next

        If iCountExeJob = 0 Then
            Procesos.EnviarActividad(log)
        End If

    End Sub

#End Region

End Module
