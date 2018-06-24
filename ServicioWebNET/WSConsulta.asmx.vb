Imports System.Data.SqlClient
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://localhost/",
                                Description:="Consultas a la Base de datos")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class WSConsulta
    Inherits System.Web.Services.WebService

    Dim conexion As SqlConnection = New SqlConnection("Data Source=DESKTOP-NLCUO3R;Initial Catalog=PRUEBA_PROPIETARIO;User ID=sa;Password=123456")
    Dim comando As SqlCommand

    <WebMethod(Description:="Consulta a la tabla Marca")>
    Public Function ConsultaMarca() As String
        Dim str As String
        conexion.Open()
        comando = New SqlCommand("Select Marca_Id, Marca_Nombre From Marca", conexion)
        Dim conjunto_datos As SqlDataReader = comando.ExecuteReader
        str = "Resumen de Marca_ID y Nombre de Marca: " & StrDup(1, vbCr)
        Do While conjunto_datos.Read
            str = str & "La Marca del ID = " & conjunto_datos.GetInt32(0) & " es " & conjunto_datos.GetString(1) & StrDup(1, vbCr)
        Loop
        conexion.Close()
        Return str
    End Function

    <WebMethod(Description:="Insertar datos en la tabla Marca")>
    Public Function InsertarMarca(Marca_Id As String, Marca_Nombre As String) As String
        conexion.Open()
        'Consulta: Insert into Marca values (Marca_Id,Marca_Nombre)
        comando = New SqlCommand("Insert into Marca values (" & CInt(Marca_Id) & ", '" & Marca_Nombre & "' )", conexion)
        comando.ExecuteNonQuery()
        conexion.Close()
        Return "Registro Creado"
    End Function


End Class