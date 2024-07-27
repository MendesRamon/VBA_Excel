Attribute VB_Name = "CONSULTANDO"
Sub ConsultarTabelaNoAccess()

    ' Definir vari�veis
    Dim conn As Object
    Dim rs As Object
    Dim strSQL As String
    Dim caminhoDoBancoDeDados As String
    Dim nomeDaTabela As String

    ' Especificar o caminho do banco de dados do Access
    caminhoDoBancoDeDados = "C:\Users\Lenovo\Desktop\ACCESS\DW.accdb"

    ' Nome da tabela que voc� deseja consultar
    nomeDaTabela = "BD_TESTE"

    ' Inicializar objetos de conex�o e registro
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    ' Estabelecer conex�o com o banco de dados do Access
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & caminhoDoBancoDeDados

    ' Construir a consulta SQL
    strSQL = "SELECT * FROM " & nomeDaTabela

    ' Executar a consulta
    rs.Open strSQL, conn
a = 0
b = a + 1
    ' Exibir os resultados (apenas um exemplo)
    Do While Not rs.EOF
        ' Fa�a algo com os dados, por exemplo, exiba-os na janela imediata
        Debug.Print rs.Fields(0).Name & ": " & rs.Fields(0).Value
         'Range("E" & b) = rs.Fields(a).Name & ": " & rs.Fields(a).Value
        ' Avan�ar para o pr�ximo registro
        rs.MoveNext
        a = a + 1
    Loop

    ' Fechar conex�o e conjunto de registros
    rs.Close
    conn.Close

    ' Liberar mem�ria
    Set rs = Nothing
    Set conn = Nothing

End Sub

