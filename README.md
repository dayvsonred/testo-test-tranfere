# testo-test-tranfere
teste




    Public Function VerificaBoletadorLiberado(ByVal CD_SISTEMA_DESTINO As String) As Boolean
        Dim Comando As New SqlCommand
        Dim Retorna As String
        
        VerificaBoletadorLiberado = False

        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "SP_MM_SEL_TB_CONTROLE_BOLET" 
                With.Parameters
                    .AddWithValue("@CD_SISTEMA_DESTINO", CD_SISTEMA_DESTINO)
                End With

                Retorna = CType(.ExecuteScalar(), String)

                If Retorna.ToString = "N" Then
                    Return True
                End If
            End With

        Finally
            Me.FecharConexao()
        End Try

    End Function

    Public Function VerificaBoletadorLiberado(ByVal CD_SISTEMA_DESTINO As String) As Boolean
        Try
            Using conexao As New SqlConnection("SuaStringDeConexão")
                conexao.Open()

                Using comando As New SqlCommand("SP_MM_SEL_TB_CONTROLE_BOLET", conexao)
                    comando.CommandType = CommandType.StoredProcedure
                    comando.Parameters.AddWithValue("@CD_SISTEMA_DESTINO", CD_SISTEMA_DESTINO)

                    Return Convert.ToString(comando.ExecuteScalar()) = "N"
                End Using
            End Using

        Catch ex As Exception
            ' Tratamento de exceções, se necessário
            Console.WriteLine("Erro no processo: " & ex.Message)
            Return False
        End Try
    End Function

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################

    Public Function ObterDataMesa(ByVal vpNM_Mesa As String) As Date
        Dim Comando As New SqlCommand
        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "DBO. FNMM_GETDATE MESA"
                With .Parameters
                    .Clear()
                    AddWithValue("@MESA", vpNM_Mesa)
                    .AddWithValue("@GETDATE", Date.Now)
                    .Add("@Return", SqlDbType.DateTime). Direction = ParameterDirection. ReturnValue
                End With
                    .ExecuteNonQuery()
                Return Framework.Convert.FromDB(.Parameters("@Return").Value, Date.Now) 
            End With

        Finally
            Me.FecharConexao()
        End Try
    End Function


    Public Function ObterDataMesa(ByVal vpNM_Mesa As String) As Date
        Try
            Using conexao As New SqlConnection("SuaStringDeConexão")
                conexao.Open()

                Using comando As New SqlCommand("DBO.FNMM_GETDATE_MESA", conexao)
                    comando.CommandType = CommandType.StoredProcedure

                    With comando.Parameters
                        .Clear()
                        .AddWithValue("@MESA", vpNM_Mesa)
                        .Add("@GETDATE", SqlDbType.DateTime).Direction = ParameterDirection.Input
                        .Parameters("@GETDATE").Value = Date.Now
                        .Add("@Return", SqlDbType.DateTime).Direction = ParameterDirection.ReturnValue
                    End With

                    comando.ExecuteNonQuery()

                    Return Convert.ToDateTime(comando.Parameters("@Return").Value)
                End Using
            End Using

        Catch ex As Exception
            ' Tratamento de exceções, se necessário
            Return Date.Now ' Retorna a data atual em caso de erro
        End Try
    End Function

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################


    Public Function VerificaTaxaCompromissada(ByVal CD EMPRESA As String, ByVal DT_REF As Date, Byval CD_TAXA As String) As Contrato. DadosTbTaxaCompromissada 
        Dim Comando As New SqlCommand
        Dim vAdapter As SqlDataAdapter
        Dim vTabela As DataTable
        Dim vResultado As New Contrato.DadosTbTaxaCompromissada

        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure .CommandText = "SP_MM_SEL_TB_TAXA_COMPROMISSADA" 
                With .Parameters
                    .Clear()
                    .AddWithValue("@CD EMPRESA", CD EMPRESA)
                    AddWithValue("@DT_REF", DT REF.ToString("yyyyMMdd")) 
                    AddWithValue("@CD_TAXA", CD_TAXA)
                End With
            End With

            vAdapter = New SqlDataAdapter(Comando)
            VTabela = New DataTable()
            VAdapter.Fill(vTabela)

            If vTabela.Rows.Count > 0 Then
                VResultado = New Contrato.DadosTbTaxaCompromissada(vTabela.Rows(0))
            End If

            Return vResultado

        Finally
            Me.FecharConexan()
        End Try

    End Function



    Public Function VerificaTaxaCompromissada(ByVal CD_EMPRESA As String, ByVal DT_REF As Date, ByVal CD_TAXA As String) As Contrato.DadosTbTaxaCompromissada
        Dim vResultado As New Contrato.DadosTbTaxaCompromissada()

        Try
            Using conexao As New SqlConnection("SuaStringDeConexão")
                conexao.Open()

                Using comando As New SqlCommand("SP_MM_SEL_TB_TAXA_COMPROMISSADA", conexao)
                    comando.CommandType = CommandType.StoredProcedure

                    comando.Parameters.AddWithValue("@CD_EMPRESA", CD_EMPRESA)
                    comando.Parameters.AddWithValue("@DT_REF", DT_REF.ToString("yyyyMMdd"))
                    comando.Parameters.AddWithValue("@CD_TAXA", CD_TAXA)

                    Using leitor As SqlDataReader = comando.ExecuteReader()
                        If leitor.HasRows Then
                            leitor.Read()
                            vResultado = New Contrato.DadosTbTaxaCompromissada(leitor)
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' Tratamento de exceções, se necessário
        End Try

        Return vResultado
    End Function

    Public Function VerificaTaxaCompromissada(ByVal CD_EMPRESA As String, ByVal DT_REF As Date, ByVal CD_TAXA As String) As List(Of TaxaCompromissada)
    Dim vResultado As New List(Of TaxaCompromissada)()

    Try
        Using conexao As New SqlConnection("SuaStringDeConexão")
            conexao.Open()

            Using comando As New SqlCommand("SP_MM_SEL_TB_TAXA_COMPROMISSADA", conexao)
                comando.CommandType = CommandType.StoredProcedure

                comando.Parameters.AddWithValue("@OCD_EMPRESA", CD_EMPRESA)
                comando.Parameters.AddWithValue("@CDT_PEF", DT_REF.ToString("yyyyMMdd"))
                comando.Parameters.AddWithValue("@OCD_TAXA", CD_TAXA)

                Using leitor As SqlDataReader = comando.ExecuteReader()
                    While leitor.Read()
                        Dim taxa As New TaxaCompromissada()
                        taxa.Usu_Valida = Convert.ToString(leitor("Usu_Valida"))
                        taxa.VL_REFERENCIAL = Convert.ToDouble(leitor("VL_REFERENCIAL"))
                        taxa.VL_DELTA = Convert.ToDouble(leitor("VL_DELTA"))
                        taxa.VL_DELTAWN = Convert.ToDouble(leitor("VL_DELTAWN"))
                        taxa.DT_REF = Convert.ToString(leitor("DT_REF"))
                        vResultado.Add(taxa)
                    End While
                End Using
            End Using
        End Using

    Catch ex As Exception
        ' Tratamento de exceções, se necessário
    End Try

    Return vResultado
End Function


Public Function VerificaTaxaCompromissada(ByVal CD_EMPRESA As String, ByVal DT_REF As Date, ByVal CD_TAXA As String) As List(Of TaxaCompromissada)
    Dim vResultado As New List(Of TaxaCompromissada)()

    Try
        Using conexao As New SqlConnection("SuaStringDeConexão")
            conexao.Open()

            Using comando As New SqlCommand("SP_MM_SEL_TB_TAXA_COMPROMISSADA", conexao)
                comando.CommandType = CommandType.StoredProcedure

                comando.Parameters.AddWithValue("@OCD_EMPRESA", CD_EMPRESA)
                comando.Parameters.AddWithValue("@CDT_PEF", DT_REF.ToString("yyyyMMdd"))
                comando.Parameters.AddWithValue("@OCD_TAXA", CD_TAXA)

                Using leitor As SqlDataReader = comando.ExecuteReader()
                    While leitor.Read()
                        Dim taxa As New TaxaCompromissada()
                        taxa.Usu_Valida = Convert.ToString(leitor("Usu_Valida"))
                        taxa.VL_REFERENCIAL = Convert.ToDouble(leitor("VL_REFERENCIAL"))
                        taxa.VL_DELTA = Convert.ToDouble(leitor("VL_DELTA"))
                        taxa.VL_DELTAWN = Convert.ToDouble(leitor("VL_DELTAWN"))
                        taxa.DT_REF = Convert.ToString(leitor("DT_REF"))
                        vResultado.Add(taxa)
                    End While
                End Using
            End Using
        End Using

    Catch ex As Exception
        ' Tratamento de exceções, se necessário
    End Try

    Return vResultado
End Function


'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################

    Public Function AtualizaStatusBoleto (ByVal SiglaSistema As String) As String 
        Dim Comando As New SqlCommand

            Try
                With Comando
                    .Connection = Me.Conexao
                .CommandText = "SP_MM_ATU_TB_STATUS_BOLET" 
                .CommandType = CommandType. StoredProcedure 
                With .Parameters
                    .Clear()
                    .AddWithValue("@DT_PROCESSO", Date.Now) 
                    .AddWithValue("@CD_BOLETADOR", SiglaSistema)
                End With

                Return Comando.ExecuteScalar().ToString

                End With
            Finally
                Me.FecharConexao()
            End Try
    End Function





    Public Function AtualizaStatusBoleto(ByVal SiglaSistema As String) As String
        Dim linhasAfetadas As Integer = 0

        Try
            Using conexao As New SqlConnection("SuaStringDeConexão")
                conexao.Open()

                Using comando As New SqlCommand("SP_MM_ATU_TB_STATUS_BOLET", conexao)
                    comando.CommandType = CommandType.StoredProcedure

                    comando.Parameters.AddWithValue("@DT_PROCESSO", Date.Now)
                    comando.Parameters.AddWithValue("@CD_BOLETADOR", SiglaSistema)

                    linhasAfetadas = comando.ExecuteNonQuery()
                End Using
            End Using

            Return linhasAfetadas.ToString()

        Catch ex As Exception
            ' Tratamento de exceções, se necessário
            Console.WriteLine("Erro no processo: " & ex.Message) 
            Return "1" ' Retorna -1 em caso de erro
        End Try
    End Function

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################


    Public Function BuscaParametrizacaoLastro(Byval CdEmpresa As String,
                                                ByVal DtMovimento As Date,
                                                Byval Sistorigem As String) As List (Of Contrato.ParametrizacaoPapeisLastro)
        Dim Comando As New SqlCommand
        Dim vAdapter As SqlDataAdapter 
        Dim vTabela As DataTable
        Dim vResultado As List (Of Contrato.ParametrizacaoPapeisLastro)

        Try
            'Executa chamada
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure
                .CommandText = "SP_MM_SEL_PARAM_LASTRO ROBO"
                With .Parameters
                    .Clear()
                    .AddWithValue("@cd_empresa", CdEmpresa)
                    .AddWithValue("@dt_movimento", DtMovimento.ToString("yyyyMMdd"))
                    .AddWithValue("@sist_origen", Sistorigen)
                End With

            End With

            vAdapter = New SqlDataAdapter(Comando)
            VTabela = New DataTable()
            VAdapter.Fill(VTabela)

            vResultado  = New List(Of Contrato.ParametrizacaoPapeisLastro)() 
            For Each vLinha As DataRow In vTabela.Rows
                vResultado.Add(New Contrato. ParametrizacaoPapeisLastro(vLinha))
            Next
            
            Return vResultado
        Finally
            Me.FecharConexan()
        End Try

    End Function






    Public Function BuscaParametrizacaoLastro(ByVal CdEmpresa As String,
                                            ByVal DtMovimento As Date,
                                            ByVal Sistorigem As String) As List(Of Contrato.ParametrizacaoPapeisLastro)
        Dim vResultado As New List(Of Contrato.ParametrizacaoPapeisLastro)()

        Try
            Using conexao As New SqlConnection("SuaStringDeConexão")
                conexao.Open()

                Using comando As New SqlCommand("SP_MM_SEL_PARAM_LASTRO ROBO", conexao)
                    comando.CommandType = CommandType.StoredProcedure

                    comando.Parameters.AddWithValue("@cd_empresa", CdEmpresa)
                    comando.Parameters.AddWithValue("@dt_movimento", DtMovimento.ToString("yyyyMMdd"))
                    comando.Parameters.AddWithValue("@sist_origen", Sistorigen)

                    Using leitor As SqlDataReader = comando.ExecuteReader()
                        While leitor.Read()
                            Dim parametrizacao As New Contrato.ParametrizacaoPapeisLastro()
                            parametrizacao.PopulateFromDataReader(leitor)
                            vResultado.Add(parametrizacao)
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' Tratamento de exceções, se necessário
        End Try

        Return vResultado
    End Function





'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################







Public Function BuscaParametrizacaoLastro(ByVal CdEmpresa As String, ByVal DtMovimento As Date, ByVal Sistorigem As String) As List(Of TaxaCompromissada)
    Dim vResultado As New List(Of TaxaCompromissada)()

    Try
        Using conexao As New SqlConnection("SuaStringDeConexão")
            conexao.Open()

            Using comando As New SqlCommand("SP_MM_SEL_PARAM_LASTRO ROBO", conexao)
                comando.CommandType = CommandType.StoredProcedure

                comando.Parameters.AddWithValue("@cd_empresa", CdEmpresa)
                comando.Parameters.AddWithValue("@dt_movimento", DtMovimento.ToString("yyyyMMdd"))
                comando.Parameters.AddWithValue("@sist_origen", Sistorigen)

                Using leitor As SqlDataReader = comando.ExecuteReader()
                    While leitor.Read()
                        Dim taxa As New TaxaCompromissada()
                        taxa.Usu_Valida = Convert.ToString(leitor("Usu_Valida"))
                        taxa.VL_REFERENCIAL = Convert.ToDouble(leitor("VL_REFERENCIAL"))
                        taxa.VL_DELTA = Convert.ToDouble(leitor("VL_DELTA"))
                        taxa.VL_DELTAWN = Convert.ToDouble(leitor("VL_DELTAWN"))
                        taxa.DT_REF = Convert.ToString(leitor("DT_REF"))
                        vResultado.Add(taxa)
                    End While
                End Using
            End Using
        End Using

    Catch ex As Exception
        ' Tratamento de exceções, se necessário
    End Try

    Return vResultado
End Function





Public Class TaxaCompromissada
    Public Property Usu_Valida As String
    Public Property VL_REFERENCIAL As Double
    Public Property VL_DELTA As Double
    Public Property VL_DELTAWN As Double
    Public Property DT_REF As String
End Class







'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################




Public Function BuscaParametrizacaoLastro(ByVal CdEmpresa As String,
                                           ByVal DtMovimento As Date,
                                           ByVal Sistorigem As String) As List(Of ParametrizacaoPapeisLastro)
    Dim vResultado As New List(Of ParametrizacaoPapeisLastro)()

    Try
        Using conexao As New SqlConnection("SuaStringDeConexão")
            conexao.Open()

            Using comando As New SqlCommand("SP_MM_SEL_PARAM_LASTRO ROBO", conexao)
                comando.CommandType = CommandType.StoredProcedure

                comando.Parameters.AddWithValue("@cd_empresa", CdEmpresa)
                comando.Parameters.AddWithValue("@dt_movimento", DtMovimento.ToString("yyyyMMdd"))
                comando.Parameters.AddWithValue("@sist_origen", Sistorigem)

                Using leitor As SqlDataReader = comando.ExecuteReader()
                    While leitor.Read()
                        Dim parametrizacao As New ParametrizacaoPapeisLastro()
                        parametrizacao.CdEmporesa = Convert.ToString(leitor("CdEmporesa"))
                        parametrizacao.IdEstoque = Convert.ToDouble(leitor("IdEstoque"))
                        parametrizacao.VlFinanc = Convert.ToDouble(leitor("VlFinanc"))
                        parametrizacao.IcOrdemConsumo = Convert.ToInt16(leitor("IcOrdemConsumo"))
                        vResultado.Add(parametrizacao)
                    End While
                End Using
            End Using
        End Using

    Catch ex As Exception
        ' Tratamento de exceções, se necessário
    End Try

    Return vResultado
End Function


'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################



Public Function SelecionarCotacoes(ByVal vpCdEmpresa As String,
                                    Optional ByVal vpNmMesa As String = "",
                                    Optional ByVal vpCdCliente As Long = 0,
                                    Optional ByVal vpIcOrigem As Long = Contrato.TipoOrigemOperacao.RoboLastro,
                                    Optional ByVal vpVLFinanc As Double = 0) As List(Of Contrato.OperacoesCompromissadas)
    Dim vResultado As New List(Of Contrato.OperacoesCompromissadas)

    Try
        Using vComando As New SqlCommand("SP_MM_SEL_COTACAO_VINCULACAO", Me.Conexao)
            vComando.CommandType = CommandType.StoredProcedure

            With vComando.Parameters
                .Clear()
                .AddWithValue("@CD_EMPRESA", vpCdEmpresa)
                If Not String.IsNullOrEmpty(vpNmMesa) Then
                    .AddWithValue("@nm_mesa", vpNmMesa)
                End If
                If vpCdCliente <> 0 Then
                    .AddWithValue("@cd_cliente", vpCdCliente)
                End If
                .AddWithValue("@ic_fonte", vpIcOrigem)
            End With

            Me.Conexao.Open()
            Using leitor As SqlDataReader = vComando.ExecuteReader()
                While leitor.Read()
                    Dim cotacao As New Contrato.OperacoesCompromissadas()
                    cotacao._nomeEmpresa = Convert.ToString(leitor("_nomeEmpresa"))
                    cotacao.codigoCotacao = Convert.ToString(leitor("codigoCotacao"))
                    cotacao._fim = Convert.ToDateTime(leitor("_fim"))
                    cotacao._NmCliente = Convert.ToString(leitor("_NmCliente"))
                    cotacao.CdCamara = Convert.ToInt32(leitor("CdCamara"))
                    cotacao._NuPrazo = Convert.ToInt32(leitor("_NuPrazo"))
                    cotacao.VlOperacao = Convert.ToDouble(leitor("VlOperacao"))
                    cotacao._CdsstatusCotacao = Convert.ToInt32(leitor("_CdsstatusCotacao"))
                    cotacao._CdIndexador = Convert.ToString(leitor("_CdIndexador"))
                    cotacao._PcIndexador = Convert.ToDouble(leitor("_PcIndexador"))
                    cotacao.CdTipoTaxa = Convert.ToInt32(leitor("CdTipoTaxa"))
                    cotacao.VLTaxaOver = Convert.ToDouble(leitor("VLTaxaOver"))
                    cotacao._VLTaxa = Convert.ToDouble(leitor("_VLTaxa"))
                    cotacao._DtInicio = Convert.ToDateTime(leitor("_DtInicio"))
                    cotacao._CdAgrupPapel = Convert.ToString(leitor("_CdAgrupPapel"))
                    cotacao.CdOperadorCotacao = Convert.ToString(leitor("CdOperadorCotacao"))
                    cotacao._DtUltAlt = Convert.ToDateTime(leitor("_DtUltAlt"))
                    cotacao.CdCliente = Convert.ToInt32(leitor("CdCliente"))
                    cotacao.CdBanco = Convert.ToString(leitor("CdBanco"))
                    cotacao.NuAgencia = Convert.ToString(leitor("NuAgencia"))
                    cotacao.NuConta = Convert.ToString(leitor("NuConta"))
                    cotacao.CdFormaLiquid = Convert.ToInt32(leitor("CdFormaLiquid"))
                    cotacao.CdOperadorAlt = Convert.ToString(leitor("CdOperadorAlt"))
                    cotacao.NuPrazoDU = Convert.ToInt32(leitor("NuPrazoDU"))

                    vResultado.Add(cotacao)
                End While
            End Using
        End Using
    Finally
        Me.FecharConexao()
    End Try

    Return vResultado
End Function




'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################

 

    Public Function sBloquearDesbloquearCliente (Optional ByVal vpUsuarioControlM As String = "",
                                                Optional Byval vpCdCliente As Long = 0,
                                                Optional ByVal vpNmCliente As String = "",
                                                Optional ByVal vpTipoAcao As Contrato.BloquearDesbloquearCliente = 0, 
                                                Optional ByRef vpMsg As String = "") As Boolean
        Dim vComando As New SqlCommand

        Try
            With vComando
                CommandTimeout = 120
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "SP_MM_BLOQUEAR_DESBLOQUEAR_CLIENTE" 
                    With .Parameters
                    AddWithValue("@cd_operador", vpUsuarioControlM) 
                    AddWithValue("@ic_funcao", vpTipoAcao)
                    If vpCdCliente <> 0 Then
                        .AddWithValue("@cd_cliente", vpCdCliente)
                    End If
                    If vpNmCliente <> "" Then
                        AddWithValue("@nm_cliente", VPN Cliente)
                    End If
                    .Add("@mensagem", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
                End With

                .ExecuteNonQuery()
            End With
            
            SBloquearDesbloquearCliente = True
            
            If vpTipoAcao = Contrato.BloquearDesbloquearCliente.bdcBloquear Then 
                If CType(vComando.Parameters("Quensagem").Value, String) <> "" Then 
                    vpMsg = CType(vComando.Parameters("Quensagem").Value, String) 
                    SBLoquearDesbloquearCliente = False
                End If
            End If

        Finally
            Me.FecharConexao()
        End Try

    End Function




Public Function sBloquearDesbloquearCliente(Optional ByVal vpUsuarioControlM As String = "",
                                             Optional ByVal vpCdCliente As Long = 0,
                                             Optional ByVal vpNmCliente As String = "",
                                             Optional ByVal vpTipoAcao As Contrato.BloquearDesbloquearCliente = 0,
                                             ByRef vpMsg As String = "") As Boolean
    Dim vSucesso As Boolean = True

    Try
        Using Me.Conexao
            Dim vComando As New SqlCommand("SP_MM_BLOQUEAR_DESBLOQUEAR_CLIENTE", Conexao)
                vComando.CommandType = CommandType.StoredProcedure
                vComando.CommandTimeout = 120

            With vComando.Parameters
                .Clear()
                .AddWithValue("@cd_operador", vpUsuarioControlM)
                .AddWithValue("@ic_funcao", vpTipoAcao)
                If vpCdCliente <> 0 Then
                    .AddWithValue("@cd_cliente", vpCdCliente)
                End If
                If Not String.IsNullOrEmpty(vpNmCliente) Then
                    .AddWithValue("@nm_cliente", vpNmCliente)
                End If
                    .Add("@mensagem", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            End With
            vComando.ExecuteNonQuery()

            If vpTipoAcao = Contrato.BloquearDesbloquearCliente.bdcBloquear AndAlso Not IsDBNull(vComando.Parameters("@mensagem").Value) Then
                    vpMsg = Convert.ToString(vComando.Parameters("@mensagem").Value)
                    vSucesso = False
            End If
        End Using

    Catch ex As Exception
        ' Tratamento de exceções, se necessário
        Console.WriteLine("Erro no processo sBloquearDesbloquearCliente: " & ex.Message) 
    End Try

    Return vSucesso
End Function







'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
















