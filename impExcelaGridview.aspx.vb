Imports System.Data.SqlClient
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports System.IO

Public Class impExcelaGridview
    Inherits System.Web.UI.Page
    Private Cnn As New SqlConnection(System.Web.Configuration.WebConfigurationManager.AppSettings("Conecta").ToString)
    Protected Sub Button_visualizar_Click(sender As Object, e As EventArgs) Handles Button_visualizar.Click
        If (FileUpload_archivo.HasFile) Then
            Try
                Dim salvar_ruta As String = "/importar_doc/" + Me.FileUpload_archivo.FileName 'Ruta para copiar el archivo
                'En caso de tener acentos el archivo, los cambia para no tener problemas al leeer los datos
                salvar_ruta = salvar_ruta.Trim.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u")
                Me.FileUpload_archivo.SaveAs(Server.MapPath(salvar_ruta)) 'Salva el archivo en la ruta importar_doc
                Dim vRuta As String = Server.MapPath(salvar_ruta) 'Obtiene la ruta completa del archivo
                Dim dt As DataTable = LeerExcel(vRuta)
                GridView1.DataSource = dt
                GridView1.DataBind()
                Me.Label_aviso.Text = "Revisa la información y presiona *Realizar importación*"
                Me.Label_aviso.Attributes.Add("style", "color:green")
            Catch ex As Exception
                Me.Label_aviso.Text = ex.Message
                Me.Label_aviso.Attributes.Add("style", "color:red")
            End Try
        Else
            Me.Label_aviso.Text = "Debes selecciona un archivo"
            Me.Label_aviso.Attributes.Add("style", "color:red")
        End If
    End Sub
    Private Function LeerExcel(ByVal filePath As String) As DataTable
        Dim dt As DataTable = New DataTable()
        Using vfile As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            Dim workbook As IWorkbook = New XSSFWorkbook(vfile)
            Dim sheet As ISheet = workbook.GetSheetAt(0)
            Dim renglon As Integer
            Dim columna As Integer
            Dim vpasa As Boolean = False
            For renglon = 0 To sheet.LastRowNum
                Dim excelRow As IRow = sheet.GetRow(renglon) 'Obtiene los renglones del archivo
                Dim dr As DataRow = dt.NewRow()
                For columna = 0 To excelRow.LastCellNum - 1
                    Dim cell As ICell = excelRow.GetCell(columna) 'Obtiene el valor de la columna
                    If cell IsNot Nothing Then
                        If vpasa = False Then
                            dt.Columns.Add(cell.ToString) 'Toma el 1er. renglón del archivo y crea los encabezados del gridview
                        Else
                            dr(columna) = cell.ToString() 'Agrega el valor de la columna
                        End If
                    End If
                Next
                If vpasa = True Then
                    dt.Rows.Add(dr) 'Agrega el registro en el gridview
                End If
                vpasa = True
            Next
        End Using
        Return dt
    End Function
    Protected Sub Button_limpiar_Click(sender As Object, e As EventArgs) Handles Button_limpiar.Click
        Me.Label_aviso.Text = ""
        Me.GridView1.DataSourceID = Nothing
        Me.GridView1.DataSource = Nothing
        Me.GridView1.DataBind()
    End Sub
    Protected Sub Button_importar_Click(sender As Object, e As EventArgs) Handles Button_importar.Click
        If Me.GridView1.Rows.Count = 0 Then
            Me.Label_aviso.Text = "No existen registros a importar"
            Me.Label_aviso.Attributes.Add("style", "color:red")
            Exit Sub
        End If
        Try
            For i As Integer = 0 To Me.GridView1.Rows.Count - 1
                Dim cSentencia = "INSERT INTO directorio " +
                "(folio,empresa,nombre,puesto,sexo)" +
                " VALUES " +
                "(@folio_v,@empresa_v,@nombre_v,@puesto_v,@sexo_v)"
                Dim command = New SqlCommand(cSentencia, Cnn)
                command.Parameters.AddWithValue("@folio_v", Me.GridView1.Rows(i).Cells(0).Text)
                command.Parameters.AddWithValue("@empresa_v", Me.GridView1.Rows(i).Cells(1).Text)
                command.Parameters.AddWithValue("@nombre_v", Me.GridView1.Rows(i).Cells(2).Text)
                command.Parameters.AddWithValue("@puesto_v", Me.GridView1.Rows(i).Cells(3).Text)
                command.Parameters.AddWithValue("@sexo_v", If(Me.GridView1.Rows(i).Cells(4).Text = "F", "Femenino", "Masculino"))
                Cnn.Open()
                command.ExecuteNonQuery()
                Cnn.Close()
            Next
            Me.Label_aviso.Text = "Se ha realizado la importación a la tabla Directorio"
            Me.Label_aviso.Attributes.Add("style", "color:green")
        Catch ex As Exception
            Me.Label_aviso.Text = ex.Message
            Me.Label_aviso.Attributes.Add("style", "color:red")
            Exit Sub
        End Try
    End Sub
End Class
