Imports System.Data.SqlClient

Public Class Procesos

    Public Shared Sub EnviarActividad(ByRef log As EXO_Log.EXO_Log)
        Dim sError As String = ""
        Dim sMensaje = ""
        Dim oDBSAP As SqlConnection = Nothing
        Dim oDtClientes As System.Data.DataTable = Nothing

        Try
            log.escribeMensaje("Iniciando Actividades sincro", EXO_Log.EXO_Log.Tipo.informacion)
            Conexiones.Connect_SQLServer(oDBSAP, "SQLSAP")
            'oDtClientes = Conexiones.GetValueDB(oDBSAP, "OCRD INNER JOIN OHEM ON OHEM.salesPrson = OCRD.SlpCode", "OCRD.CardCode, OCRD.CardName, OHEM.empID, OCRD.U_EXO_EnvioDia", "OCRD.U_EXO_EnvioDia LIKE '%" & Weekday(Now, FirstDayOfWeek.Monday) & "%'")

            'Dim sql As String = "select OCRD.CardCode, OCRD.CardName, OHEM.empID, OCRD.U_EXO_EnvioDia " +
            '    " from OCRD  INNER JOIN OHEM  ON OHEM.salesPrson = OCRD.SlpCode " +
            '    " where OCRD.U_EXO_EnvioDia LIKE '%" & Weekday(Now, FirstDayOfWeek.Monday) & "%'"
            'Conexiones.FillDtDB(oDBSAP, oDtClientes, sql)

            Dim sSQL As String = "SELECT OCRD.CardCode, OCRD.CardName, OHEM.empID, OCRD.U_EXO_EnvioDia FROM OCRD INNER JOIN OHEM ON OHEM.salesPrson = OCRD.SlpCode WHERE OCRD.U_EXO_EnvioDia LIKE '%" & Weekday(Now, FirstDayOfWeek.Monday) & "%'"
            oDtClientes = New System.Data.DataTable()
            Conexiones.FillDtDB(oDBSAP, oDtClientes, sSQL)

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            sMensaje = sError
            log.escribeMensaje("Error iniciando: " + sMensaje, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            sMensaje = sError
            log.escribeMensaje("Error iniciando: " + sMensaje, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLServer(oDBSAP)
            GenerarActividad(oDtClientes, log)
        End Try
    End Sub

    Public Shared Sub GenerarActividad(oDT As System.Data.DataTable, log As EXO_Log.EXO_Log)
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim oCompServ As SAPbobsCOM.CompanyService = Nothing
        Dim oActivityService As SAPbobsCOM.ActivitiesService = Nothing
        Dim oCLGParams As SAPbobsCOM.ActivityParams = Nothing
        Dim oOCLG As SAPbobsCOM.Activity = Nothing

        Try
            If oDT.Rows.Count > 0 Then
                Conexiones.Connect_Company(oCompany)
                oCompServ = oCompany.GetCompanyService

                oActivityService = CType(oCompServ.GetBusinessService(SAPbobsCOM.ServiceTypes.ActivitiesService), SAPbobsCOM.ActivitiesService)

                For Each row As DataRow In oDT.Rows

                    Dim a As String = row.Item("CardCode").ToString()
                Next

                For Each row As DataRow In oDT.Rows

                    'if comprobar empleado rellenado
                    If row.Item("empID").ToString <> "" And row.Item("empID") IsNot Nothing Then
                        oOCLG = CType(oActivityService.GetDataInterface(SAPbobsCOM.ActivitiesServiceDataInterfaces.asActivity), SAPbobsCOM.Activity)

                        oOCLG.Activity = SAPbobsCOM.BoActivities.cn_Task
                        oOCLG.CardCode = row.Item("CardCode").ToString
                        oOCLG.ActivityType = -1
                        oOCLG.Subject = 106
                        oOCLG.HandledByEmployee = row.Item("empID").ToString
                        oOCLG.Details = "Día de reparto."
                        oOCLG.Notes = "Hoy es día de reparto."
                        oOCLG.StartDate = Now

                        oCLGParams = oActivityService.AddActivity(oOCLG)
                        Dim ActividadCode As String = CStr(CInt(oCLGParams.ActivityCode))

                        If ActividadCode <> "" Or ActividadCode <> 0 Then
                            'Actividad creada correctamente
                            'log un ok con el código de cliente
                            log.escribeMensaje("OK: " + row.Item("CardCode").ToString, EXO_Log.EXO_Log.Tipo.informacion)
                        Else
                            'Error al crear la actividad
                            'log error  con el codiog de cliente y el mensaje de error
                            'oCompany.GetLastErrorDescription
                            log.escribeMensaje("ERROR: " + oCompany.GetLastErrorDescription.ToString, EXO_Log.EXO_Log.Tipo.error)
                        End If
                    Else
                        'else mensaje en el log
                        log.escribeMensaje("El cliennte " + row.Item("CardCode") + " no tiene un empleado asociado!", EXO_Log.EXO_Log.Tipo.advertencia)
                    End If

                Next

            End If
        Catch ex As Exception
            'log error
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oCompServ IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompServ)
                oCompServ = Nothing
            End If
            If oOCLG IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCLG)
                oOCLG = Nothing
            End If

            Conexiones.Disconnect_Company(oCompany)
        End Try

    End Sub
End Class
