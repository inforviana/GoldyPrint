Imports System
Imports System.IO
Imports System.IO.Ports
Imports MySql.Data
Imports MySql.Data.MySqlClient


Public Class Form1
    Private Property SerialPort As SerialPort
    Dim mysqlconn As New MySqlConnection

    Dim queryCab As String
    Dim linha As String
    Dim server, user, pw, db As String


    Sub tal(texto As String) 'imprimir texto linha a linha
        SerialPort.WriteLine(texto)
    End Sub

    Sub cortar() 'cortar o papel
        SerialPort.Write(Chr(29) + Chr(86) + Chr(65) + Chr(0))
    End Sub

    Sub lerconfig()
        Dim linha() As String
        Dim conf, val As String

        If (File.Exists("config.ini")) Then 'ficheiro ja existe
            'ler o ficheiro de configuracao
            Dim ficheiro As StreamReader = New StreamReader("config.ini")
            
            While (ficheiro.Peek <> -1)
                linha = ficheiro.ReadLine.Split("=")
                conf = linha(0)
                val = linha(1)

                If (conf = "server") Then
                    server = val
                ElseIf (conf = "user") Then
                    user = val
                ElseIf (conf = "pw") Then
                    pw = val
                ElseIf (conf = "db") Then
                    db = val
                ElseIf (conf = "porta") Then
                    txtPorta.Text = val
                ElseIf (conf = "baud") Then
                    txtVelocidade.Text = val
                End If
            End While
            ficheiro.Close()
            mysqlconn.ConnectionString = "server=" + server + ";user id=" + user + ";password=" + pw + ";database=" + db + ";"
            SerialPort.PortName = txtPorta.Text
            SerialPort.BaudRate = txtVelocidade.Text
        Else 'criar o ficheiro de configuracao com as configuracoes de fabrica
            criarconfig()
        End If
    End Sub

    Sub criarconfig()
        Dim ficheiro As StreamWriter = New StreamWriter("config.ini", False)
        ficheiro.WriteLine("server=localhost")
        ficheiro.WriteLine("user=root")
        ficheiro.WriteLine("pw=")
        ficheiro.WriteLine("db=goldylocks")
        ficheiro.WriteLine("porta=COM9")
        ficheiro.WriteLine("baud=9600")
        ficheiro.Close()
        lerconfig()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        queryCab = "select spooler.ean, spooler.qtd, clientes.nome, encomenda.data_criacao, time(encomenda.data_criacao), clientes.telemovel, clientes.telefone, equipamentos.observacoes_equipamento, encomenda.observacoes_encomenda, encomenda.codigo_unico, clientes.nif, tipos_documento.tipo_documento, encomenda.numero_documento, clientes.morada, clientes.cp from spooler left join encomenda on encomenda.id_encomenda = spooler.qtd left join clientes on clientes.cod_sage = encomenda.id_cliente left join equipamentos on equipamentos.id_equipamento = encomenda.id_equipamento left join tipos_documento on tipos_documento.id_tipo_documento = encomenda.id_tipo_documento where spooler.ean = 'NR'"
        linha = "******************************************"
        SerialPort = New SerialPort
        

        'mysql

        lerconfig()
    End Sub

    Private Sub cmdTeste_Click(sender As Object, e As EventArgs) Handles cmdTeste.Click
        SerialPort.PortName = txtPorta.Text
        SerialPort.BaudRate = txtVelocidade.Text

        SerialPort.Open()
        SerialPort.WriteLine(linha) 'largura epson tm
        SerialPort.Write(Chr(29) + Chr(86) + Chr(65) + Chr(0)) 'corte
        SerialPort.Close()
    End Sub

    Private Sub timer_Tick(sender As Object, e As EventArgs) Handles timer.Tick
        mysqlconn.Open()
        Dim mysqlcmd As New MySqlCommand
        Dim mysqlcmd2 As New MySqlCommand
        Dim sqlLinhas As New MySqlCommand
        Dim mysqldatareader As MySqlDataReader
        Dim drLinhas As MySqlDataReader
        Dim documentos As New ArrayList
        Dim opcoes As New ArrayList
        Dim conf, valor As String
        Dim nome, morada, localidade, nif, telefone, telemovel, email, site As String
        Dim total, totalDoc As Double
        Dim nomeArtigo As String


        mysqlcmd.Connection = mysqlconn
        mysqlcmd2.Connection = mysqlconn
        sqlLinhas.Connection = mysqlconn

        mysqlcmd.CommandText = queryCab
        mysqlcmd2.CommandText = "delete from spooler"

        mysqldatareader = mysqlcmd.ExecuteReader 'obter os documentos a imprimir
        While (mysqldatareader.Read)
            Dim dict As New Dictionary(Of Integer, Object)

            For i As Integer = 0 To (mysqldatareader.FieldCount - 1)
                dict.Add(i, mysqldatareader(i))
            Next
            documentos.Add(dict)
        End While
        mysqldatareader.Close()

        mysqlcmd.CommandText = "select * from config" 'obter dados da empresa
        mysqldatareader = mysqlcmd.ExecuteReader
        While (mysqldatareader.Read)
            conf = mysqldatareader.GetString(1)
            valor = mysqldatareader.GetString(2)

            If (conf = "nome_empresa") Then
                nome = valor
            ElseIf (conf = "morada_empresa") Then
                morada = valor
            ElseIf (conf = "localidade_empresa") Then
                localidade = valor
            ElseIf (conf = "nif_empresa") Then
                nif = valor
            ElseIf (conf = "telefone_empresa") Then
                telefone = valor
            ElseIf (conf = "telemovel_empresa") Then
                telemovel = valor
            ElseIf (conf = "email_empresa") Then
                email = valor
            ElseIf (conf = "site_empresa") Then
                site = valor
            End If
        End While
        mysqldatareader.Close()


        SerialPort.Open() 'abrir a porta de serie a utilizar
        SerialPort.Write(Chr(27) + Chr(116) + Chr(3)) 'CODEPAGE DE PORTUGAL

        For Each dat As Dictionary(Of Integer, Object) In documentos
            sqlLinhas.CommandText = "select mov_encomendas.id_artigo, artigos.nome, mov_encomendas.observacao, mov_encomendas.quantidade, mov_encomendas.preco from mov_encomendas left join artigos on artigos.cod_barras = mov_encomendas.id_artigo where id_encomenda=" + dat(1).ToString
            If (nome.Length > 0) Then tal(nome)
            If (morada.Length > 0) Then tal(morada)
            If (localidade.Length > 0) Then tal(localidade)
            If (nif.Length > 0) Then tal("NIF: " + nif)
            If (telefone.Length > 0) Then tal("Telf. " + telefone)
            If (telemovel.Length > 0) Then tal("Telem." + telemovel)
            If (email.Length > 0) Then tal(email)
            If (site.Length > 0) Then tal(site)
            tal(linha)
            '   "******************************************" REGUA
            tal("         " + dat(11).ToString + " N. " + dat(12).ToString + "             ")
            tal(linha)
            tal("Cliente: " + dat(2).ToString)
            If (dat(10).ToString.Length > 0) Then
                tal("Contribuinte: " + dat(10).ToString)
            Else
                tal("Contribuinte: Consumidor Final")
            End If

            tal("Morada: " + dat(13).ToString)
            tal("Localidade: " + dat(14).ToString)
            tal("Telefone: " + dat(6).ToString)
            tal("Telemovel: " + dat(5).ToString)
            tal("Data/Hora: " + dat(3).ToString)
            tal("------------------------------------------")
            If (dat(7).ToString.Length > 0) Then tal("Equipamento: " + dat(7).ToString)
            If (dat(8).ToString.Length > 0) Then tal("Acessorios: " + dat(8).ToString)
            tal("------------------------------------------")
            tal("Qtd    Artigo      P.Unit.        Total")
            tal("------------------------------------------")
            drLinhas = sqlLinhas.ExecuteReader
            totalDoc = 0
            While (drLinhas.Read)
                If (drLinhas.GetString(1).Length > 10) Then 'limitar o numero maximo de caracteres nos nomes dos artigos
                    nomeArtigo = drLinhas.GetString(1).Substring(0, 9)
                Else
                    nomeArtigo = drLinhas.GetString(1)
                End If
                total = drLinhas.GetInt32(3) * drLinhas.GetDouble(4) 'total de cada linha
                totalDoc += total 'total do documento
                tal(drLinhas.GetInt32(3).ToString + "  x  " + nomeArtigo + "     " + drLinhas.GetDouble(4).ToString + " eur    " + total.ToString + " eur") 'linha 
                If (drLinhas.GetString(2).Length > 0) Then tal("      " + drLinhas.GetString(2)) 'observacoes
            End While
            drLinhas.Close()
            tal("")
            tal("TOTAL A PAGAR: " + totalDoc.ToString + " euros")
            tal("")
            tal("")
            tal("***        CODIGO WEB: " + dat(9).ToString + "       ***")
            'TODO: moradas para as guias de transporte
            tal("*************** GOLDYLOCKS *************** ")
            cortar()
        Next



        mysqlcmd2.ExecuteNonQuery() 'limpar o spooler de impressao

        SerialPort.Close() 'fechar a porta de serie

        mysqlconn.Close() 'fechar a ligacao ao mysql
    End Sub
End Class
