Attribute VB_Name = "BANCO_DADOS"
Sub InserirDadosNaTabelaAccess()

    ' Definir vari�veis
    Dim conn                        As Object
    Dim strSQL                      As String
    Dim caminhoDoBancoDeDados       As String
    Dim nomeDaTabela                As String
    Dim valorCampo1                 As String
    Dim valorCampo2                 As Integer  ' Exemplo de tipo de dados diferente
    Dim vlr1                        As String
    Dim vlr2                        As Double

    Set tabela = Range("A2:A16")

    ' Especificar o caminho do banco de dados do Access
    caminhoDoBancoDeDados = "C:\Users\Lenovo\Desktop\PROJETOS_AUTOMACAO\CONSOLIDANDO_ARQUIVOS\BANCO_DADOS\BANCO_DE_DADOS.accdb"

    ' Nome da tabela que voc� deseja atualizar
    nomeDaTabela = "tabela_teste"

    ' Valores que voc� deseja inserir
   ' valorCampo1 = "A"
    'valorCampo2 = 123  ' Exemplo de valor inteiro

    ' Inicializar objeto de conex�o
    Set conn = CreateObject("ADODB.Connection")

    ' Estabelecer conex�o com o banco de dados do Access
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & caminhoDoBancoDeDados

    For Each Rng In tabela
    
    'atribuindo valor din�mico a vari�veis
        valorCampo1 = Rng
        valorCampo2 = Rng.Offset(, 1)
    
    
        ' Construir a consulta SQL para inserir dados
        strSQL = "INSERT INTO " & nomeDaTabela & " (LETRA, NUMERO) VALUES ('" & valorCampo1 & "', " & valorCampo2 & ")"
    
        ' Executar a consulta
        conn.Execute strSQL
        
    Next

    ' Fechar a conex�o
    conn.Close

    ' Liberar mem�ria
    Set conn = Nothing
    
    MsgBox "Dados Inseridos com Sucesso!", vbInformation, "Insert em BD"
    
End Sub

