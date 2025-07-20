Attribute VB_Name = "mdlCargar"
Public Conex As ADODB.Connection
Public Recor As ADODB.Recordset
Public BD As String
Public Cnn As Connection
Public Rst As Recordset
Public Num As String

Public Sub Cargar()
'Abrir conexiones con las tablas de la base de datos
    BD = "C:\Archivos de programa\Sistema de Pasantías\Taller.mdb" 'App.Path & "\Tickets.mdb"
    Set Conex = New ADODB.Connection
    Set Recor = New ADODB.Recordset
    Conex.Open "provider=microsoft.jet.oledb.4.0; data source=" & BD
        DEPago_de_Taller.Connection1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & BD & ";Persist Security Info=False"
End Sub
