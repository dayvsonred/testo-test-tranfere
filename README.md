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



Public Function TitulosParaLastro(ByVal vpIdEstoque As String _,
        ByVal vpNmMesa As String _,
        ByVal vpCdEmpresa As String _,
        Byval vpdtInicio As Date _,
        Byval vpdtFim As Date _,
        Optional Byval vpCdTitulo As String = "" _,
        Optional ByVal vpCdCamara As Long = 0 _,
        Optional ByVal vpAutomatico As Integer = Contrato.FormaLastro.flRobo _,
        Optional Byval vpOrigem As Integer = Contrato.TipoOrigemOperacao.Integrador _,
        Optional ByVal vpSistOrigem As String = "") As List(of Contrato.PapeisUsadosParaLastro)

        Dim vComando As New SqlCommand
        Dim retorno As New DataTable()

        Dim vAdapter As SqlDataAdapter 
        Dim vTabela As DataTable
        Dim vResultado As List(Of Contrato.PapeisUsadosParaLastro)

        Try
            With vComando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure
                If vpIdEstoque = "T" Then           '-Estoque Terceiros
                    .CommandText = "SP_MM_SEL TERCEIROS LASTRO_ROBO"
                ElseIf vpIdEstoque = "p" Then       '-Estoque Proprio
                    .CommandText = "SP_MM_SEL PROPRIO LASTRO_ROBO"
                ElseIf vpIdEstoque = "L" Then       '-Estoque LM
                    .CommandText = "SP_MM_SEL_LM_LASTRO ROBO"
                End If

                With .Parameters
                    Clear()
                    AddWithValue("@NM_MESA", vpNmMesa)
                    AddWithValue("@CD_EMPRESA", vpCdEmpresa)
                    AddWithValue("@DT_INICIO", vpdtInicio.ToString("yyyyMMdd")) 
                    AddWithValue("@DT_FIM", vpdtFim.ToString("yyyyMMdd"))


                    If vpCdTitulo <> "" Then
                        .AddWithValue("@CD_TITULO", vpCdTitulo)
                    End If
                    If vpCdCamara <> 0 Then
                        .AddWithValue("@CD_CAMARA", vpCdCamara)
                    End If
                        .AddWithValue("@IC_AUTOMATICO", vpAutomatico) 
                        .AddWithValue("@IC_ORIGEM", vpOrigem) 
                        .AddWithValue("@SIST_ORIGEM", vpSistOrigem)
                    End With
                End With

                vAdapter = New SqlDataAdapter(vComando)
                vTabela = New DataTable()
                vAdapter.Fill(vTabela)

                vResultado = New List(Of Contrato. PapeisUsadosParaLastro)() 
                For Each vLinha As DataRow In vTabela.Rows
                    vResultado.Add(New Contrato. PapeisUsadosParaLastro(vLinha))
                Next

                Return vResultado

            Finally
                Me.FecharConexao()
            End Try
            
    End Function



    Public Function TitulosParaLastro(ByVal vpIdEstoque As String,
                                    ByVal vpNmMesa As String,
                                    ByVal vpCdEmpresa As String,
                                    ByVal vpdtInicio As Date,
                                    ByVal vpdtFim As Date,
                                    Optional ByVal vpCdTitulo As String = "",
                                    Optional ByVal vpCdCamara As Long = 0,
                                    Optional ByVal vpAutomatico As Integer = Contrato.FormaLastro.flRobo,
                                    Optional ByVal vpOrigem As Integer = Contrato.TipoOrigemOperacao.Integrador,
                                    Optional ByVal vpSistOrigem As String = "") As List(Of Contrato.PapeisUsadosParaLastro)

        Dim vResultado As New List(Of Contrato.PapeisUsadosParaLastro)

        Try
            Using vComando As New SqlCommand()
                vComando.Connection = Me.Conexao
                vComando.CommandType = CommandType.StoredProcedure

                Select Case vpIdEstoque
                    Case "T"
                        vComando.CommandText = "SP_MM_SEL_TERCEIROS_LASTRO_ROBO"
                    Case "P"
                        vComando.CommandText = "SP_MM_SEL_PROPRIO_LASTRO_ROBO"
                    Case "L"
                        vComando.CommandText = "SP_MM_SEL_LM_LASTRO_ROBO"
                End Select

                With vComando.Parameters
                    .AddWithValue("@NM_MESA", vpNmMesa)
                    .AddWithValue("@CD_EMPRESA", vpCdEmpresa)
                    .AddWithValue("@DT_INICIO", vpdtInicio.ToString("yyyyMMdd"))
                    .AddWithValue("@DT_FIM", vpdtFim.ToString("yyyyMMdd"))

                    If Not String.IsNullOrEmpty(vpCdTitulo) Then
                        .AddWithValue("@CD_TITULO", vpCdTitulo)
                    End If

                    If vpCdCamara <> 0 Then
                        .AddWithValue("@CD_CAMARA", vpCdCamara)
                    End If

                    .AddWithValue("@IC_AUTOMATICO", vpAutomatico)
                    .AddWithValue("@IC_ORIGEM", vpOrigem)
                    .AddWithValue("@SIST_ORIGEM", vpSistOrigem)
                End With

                Using leitor As SqlDataReader = vComando.ExecuteReader()
                    While leitor.Read()
                        Dim papeis As New Contrato.PapeisUsadosParaLastro()

                            papeis.CD_TITULO = Convert.ToString(leitor("CD_TITULO"))
                            papeis.DT_Vencimento = Convert.ToDateTime(leitor("DT_Vencimento"))
                            papeis.DT_Vencimento_Operacao = Convert.ToString(leitor("DT_Vencimento_Operacao"))
                            papeis.CD_BACEN = Convert.ToString(leitor("CD_BACEN"))
                            papeis.Disponivel = Convert.ToString(leitor("Disponivel"))
                            papeis.Categoria = Convert.ToString(leitor("Categoria"))
                            papeis.Vl_Pu_550 = Convert.ToString(leitor("Vl_Pu_550"))
                            papeis.DT_Posicao = Convert.ToString(leitor("DT_Posicao"))
                            papeis.IC_PRIORIDADE = Convert.ToString(leitor("IC_PRIORIDADE"))
                            papeis.Cd_Papel = Convert.ToString(leitor("Cd_Papel"))
                            papeis.Ds_Papel = Convert.ToString(leitor("Ds_Papel"))
                            papeis.Cd_ETQ = Convert.ToString(leitor("Cd_ETQ"))
                            papeis.Cd_Agrup_Papel = Convert.ToString(leitor("Cd_Agrup_Papel"))
                            papeis.IC_EXCECOES = Convert.ToString(leitor("IC_EXCECOES"))
                            papeis.QTDE_RESERVADA = Convert.ToString(leitor("QTDE_RESERVADA"))
                            papeis.Valor = Convert.ToString(leitor("Valor"))
                            papeis.Vl_Pu_Volta = Convert.ToString(leitor("Vl_Pu_Volta"))
                            papeis.TipoEstoque = Convert.ToString(leitor("TipoEstoque"))
                            papeis.QTDE_Titulo = Convert.ToString(leitor("QTDE_Titulo"))
                            papeis.QTDE_Utilizada = Convert.ToString(leitor("QTDE_Utilizada"))
                            papeis.QTDE_Disponivel = Convert.ToString(leitor("QTDE_Disponivel"))
                            papeis.COD_SRIE = Convert.ToString(leitor("COD_SRIE"))
                            papeis.VL_Pu_Mercado = Convert.ToString(leitor("VL_Pu_Mercado"))
                            papeis.TipoPrecoUnitario = Convert.ToString(leitor("TipoPrecoUnitario"))
                            papeis.TipoPU_Operacao = Convert.ToString(leitor("TipoPU_Operacao"))

                        vResultado.Add(papeis)
                    End While
                End Using
            End Using
        Finally
            Me.FecharConexao()
        End Try

        Return vResultado
    End Function






Public Class PapeisUsadosParaLastro 
        Inherits EntidadeBase

    Public Property CD_TITULO As String 
    Public Property DT_Vencimento As Date 
    Public Property DT_Vencimento_Operacao As String 
    Public Property CD_BACEN As String 
    Public Property Disponivel As String 
    Public Property Categoria As String 
    Public Property Vl_Pu_550 As String 
    Public Property DT_Posicao As String 
    Public Property IC_PRIORIDADE As String 
    Public Property Cd_Papel As String 
    Public Property Ds_Papel As String 
    Public Property Cd_ETQ As String 
    Public Property Cd_Agrup_Papel As String 
    Public Property IC_EXCECOES As String 
    Public Property QTDE_RESERVADA As String 
    Public Property Valor As String 
    Public Property Vl_Pu_Volta As String 
    Public Property TipoEstoque As String 
    Public Property QTDE_Titulo As String 
    Public Property QTDE_Utilizada As String 
    Public Property QTDE_Disponivel As String 
    Public Property COD_SRIE As String 
    Public Property VL_Pu_Mercado As String 
    Public Property TipoPrecoUnitario As String 
    Public Property TipoPU_Operacao As String 

End Class

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################


Public Function ConsultaPU5550(Eyval CdPapel As Integer) _ 
                                As Contrato.DadosPU550
    Dim vComando As New SqlCommand
    Dim vAdapter As SqlDataAdapter 
    Dim vTabela As DataTable
    Dim vResultado As New Contrato.DadosPU550

    Try
        With vComando
        .Connection = Me.Conexao
        .CommandText = "SP_MM_CONSULTA PU550" 
        .CommandType = CommandType.StoredProcedure
        With .Parameters
            .Clear()
            .AddWithValue("@CD_PAPEL", CdPapel)
            End With
        End With
        
        vAdapter = New SqlDataAdapter(vComando) 
        vTabela = New DataTable() 
        vAdapter.Fill(vTabela)
        If vTabela.Rows.Count > 0 Then
            vResultado = New Contrato.DadosPUS50(vTabela.Rows(0))
        End If

        Return vResultado

    Finally
        Me.FecharConexao()
    End Try
End Function





Public Function ConsultaPU5550(ByVal CdPapel As Integer) As Contrato.DadosPU550
    Dim vComando As New SqlCommand
    Dim vResultado As New Contrato.DadosPU550

    Try
        With vComando
            .Connection = Me.Conexao
            .CommandText = "SP_MM_CONSULTA_PU550" ' Corrigido o nome do procedimento armazenado
            .CommandType = CommandType.StoredProcedure
            With .Parameters
                .Clear()
                .AddWithValue("@CD_PAPEL", CdPapel)
            End With
        End With

        Using leitor As SqlDataReader = vComando.ExecuteReader()
            If leitor.Read() Then
                vResultado.VL_PU_RET = Convert.ToString(leitor("VL_PU_RET"))
            End If
        End Using
    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro durante a consulta ConsultaPU5550: " & ex.Message)
    End Try

    Return vResultado
End Function




'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################



    Public Function ProximaSequencia(ByVal CdCotacao As String) As Integer 

        Dim Comando As New SqlCommand
        Dim Retorna As Integer

        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "SP_MM_RETORNA_NU_SEQ"

                With .Parameters
                    .Clear()
                    .AddWithValue("@CD_COTACAO", CdCotacao)
                End With

                Retorna CType(.ExecuteScalar(), Integer)

                If Retorna >= 0 Then
                    Return Retorna
                Else
                    Return -1
                End If

            End with

        Finally
            Me.FecharConexao()
        End Try

    End Function



Public Function ProximaSequencia(ByVal CdCotacao As String) As Integer
    Dim Retorna As Integer = -1

    Try
        Using Comando As New SqlCommand("SP_MM_RETORNA_NU_SEQ", Me.Conexao)
            Comando.CommandType = CommandType.StoredProcedure
            Comando.Parameters.AddWithValue("@CD_COTACAO", CdCotacao)

            Using leitor As SqlDataReader = Comando.ExecuteReader()
                If leitor.Read() Then
                    Retorna = Convert.ToInt32(leitor(0))
                End If
            End Using
        End Using
    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao retornar a próxima sequência: " & ex.Message)
    Finally
        Me.FecharConexao()
    End Try

    Return Retorna
End Function




'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################






    Public Function GeraComando (ByVal vpCdEmpresa As String,
                                    Optional ByVal vpCdCustodia As String = "5",
                                    Optional ByVal vpCdTipoCliente As String = "CM") As String

        Dim Comando As New SqlCommand
        Dim Retorna As String

        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType. StoredProcedure
                .CommandText = "SP_GERA_NUMERO_DOC"

                With Parameters
                    .Clear()
                    .AddWithValue("@CD_CUSTODIA", vpCdCustodia)
                    .AddWithValue("@CD_TIPO_CLIENTE", vpCdTipoCliente) 
                    .AddWithValue("@CD_EMPRESA", vpCdEmpresa)
                End With

                Retorna = CType(.ExecuteScalar(), String)
                
                If Retorna Is DBNull.Value Then
                    GeraComando = "-1"
                Else
                    GeraComando = Retorna
                End If

            End With

        Finally
            Me.FecharConexao()
        End Try

    End Function





    Public Function GeraComando(ByVal vpCdEmpresa As String,
                            Optional ByVal vpCdCustodia As String = "5",
                            Optional ByVal vpCdTipoCliente As String = "CM") As String

    Dim Retorna As String = "-1"

    Try
        Using Comando As New SqlCommand("SP_GERA_NUMERO_DOC", Me.Conexao)
            Comando.CommandType = CommandType.StoredProcedure

            With Comando.Parameters
                .Clear()
                .AddWithValue("@CD_CUSTODIA", vpCdCustodia)
                .AddWithValue("@CD_TIPO_CLIENTE", vpCdTipoCliente)
                .AddWithValue("@CD_EMPRESA", vpCdEmpresa)
            End With

            Using leitor As SqlDataReader = Comando.ExecuteReader()
                If leitor.Read() Then
                    If Not leitor.IsDBNull(0) Then
                        Retorna = leitor.GetString(0)
                    End If
                End If
            End Using
        End Using
    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao gerar o comando: " & ex.Message)
    Finally
        Me.FecharConexao()
    End Try

    Return Retorna
End Function








'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################





Public Function IdentificarCotacao(ByVal CdCotacao As String) _ As Contrato.IdentificarCotacao
    Dim vComando As New SqlCommand
    Dim vAdapter As SqlDataAdapter
    Dim vTabela As DataTable
    Dim vResultado As New Contrato.IdentificarCotacao

    Try
    With vComando
        .Connection = Me.Conexao
        .CommandType = CommandType.StoredProcedure 
        .CommandText = "SP_MM_SEL_COTACAO"

        With .Parameters
        .Clear()
        .AddWithValue("@cd_cotacao", CdCotacao)
        End With
    End With

    vAdapter = New SqlDataAdapter(vComando) 
    vTabela = New DataTable() 
    vAdapter.Fill(vTabela)

    If vTabela.Rows.Count > 0 Then
        vResultado = New Contrato.IdentificarCotacao (vTabela.Rows(0))
    End If

    Return vResultado

    Catch ex As SqlException 
        Return Nothing
    Finally
        Me.FecharConexao()
    End Try
End Function


Public Function IdentificarCotacao(ByVal CdCotacao As String) As Contrato.IdentificarCotacao
    Dim vResultado As New Contrato.IdentificarCotacao

    Try
        Using vComando As New SqlCommand("SP_MM_SEL_COTACAO", Me.Conexao)
            vComando.CommandType = CommandType.StoredProcedure
            vComando.Parameters.AddWithValue("@cd_cotacao", CdCotacao)

            Using leitor As SqlDataReader = vComando.ExecuteReader()
                If leitor.Read() Then
                    vResultado = New Contrato.IdentificarCotacao(leitor)
                End If
            End Using
        End Using
    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao identificar a cotação: " & ex.Message)
    Finally
        Me.FecharConexao()
    End Try

    Return vResultado
End Function



'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################


Public Function GravaLogOperRobo(@yval vIdentCot As Contrato.IdentificarCotacao, 
                                    ByVal vNu_Max_Seq As Integer,
                                    Byval vListIdentPapeis As List(of Contrato. PapeisUsadosParaLastro)) As Boolean

        Dim Comando As New SqlCommand
        Dim Retorna As String
        Dim y As Integer


        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure
                .CommandText = "SP_MM_INC_TB_LOG_OPER_ROBO"
            End With

            For y = 0 To vListIdentPapeis.Count - 1

                If vListIdentPapeis(y).CD_TITULO <> "" Then

                    With Comando

                        vNu_Max_Seq = vNu_Max_Seq + 1

                        With Parameters
                            .Clear()
                            .AddWithValue("CDT MOVIMENTO", Date.Now)
                            .AddWithValue("@CD_COTACAO", vIdentCot.CD_COTACAO)
                            .AddWithValue("@NU_SEQ", vNu_Max_Seq)
                            .AddWithValue("@CD_OPERACAO_PROD", vIdentCot.CD_OPERACAO_PROD)
                            .AddWithValue("CDT_INICIO", vIdentCot.DT_INICIO)
                            .AddWithValue("CDT_FIM", vIdentCot.DT_FIM)
                            .AddWithValue("ENU_PRAZO DU", vIdentCot.NU_PRAZO_DU)
                            .AddWithValue("@VL TAXA", vIdentCot.VL_TAXA)
                            .AddWithValue("@VL_OPERACAO", vIdentCot.VL_OPERACAO)
                            .AddWithValue("@CD_INDEXADOR", vIdentCot.CD_INDEXADOR)
                            .AddWithValue("@CD_OPERADOR_COTACAO", vIdentCot.CD_OPERADOR_COTACAO)
                            .AddWithValue("@CD_CAIXA", vListIdentPapeis(y).TipoEstoque)
                            .AddWithValue("@CD_PAPEL", vListIdentPapeis(y).Cd_Papel)
                            .AddWithValue("@CD_TITULO", vListIdentPapeis(y).CD_TITULO) 
                            .AddWithValue("@CD_BACEN", vListIdentPapeis(y).CD_BACEN) 
                            .AddWithValue("@OT_VENCIMENTO", vListIdentPapeis(y).DT_Vencimento)
                            .AddWithValue("@DT VENCIMENTO_OPERACAO", vListIdentPapeis(y).DT_Vencimento_Operacao) 
                            .AddWithValue("@CD_ETQ", vListIdentPapeis(y).cd_ETQ)
                            If UCase(vListIdentPapeis (y). TipoPrecoUnitario) = "PUMERCADO" Then 
                                .AddWithValue("@VL_PU", vListIdentPapeis(y).Vl_Pu_Mercado)
                            Else
                                .AddWithValue("@VL_PU", vListIdentPapeis(y).Vl_Pu_550)
                            End If
                            .AddWithValue("@VL PU_550", vListIdentPapeis(y).Vl_Pu_Volta)
                            .AddWithValue("@NU_QUANTIDADE TITULO", vListIdentPapeis(y).QTDE_Titulo)
                            .AddWithValue("NU QUANTIDADE UTILIZADA", vListIdentPapeis(y).QTDE_Utilizada) 
                            .AddWithValue("ONU QUANTIDADE DISPONIVEL", vListIdentPapeis(y).Disponivel) 
                            .AddWithValue("@DS_ERRO", "")
                        End With
        
                        Retorna CType(.ExecuteNonQuery(), String)

                        If Retorna = "0" Then
                            Return False
                        End If
                    End With
                End If

            Next

            Return True

        Finally
            Me.FecharConexao()
        End Try

End Function








Public Function GravaLogOperRobo(ByVal vIdentCot As Contrato.IdentificarCotacao,
                                  ByVal vNu_Max_Seq As Integer,
                                  ByVal vListIdentPapeis As List(Of Contrato.PapeisUsadosParaLastro)) As Boolean

    Dim Comando As New SqlCommand

    Try
        Using Comando
            Comando.Connection = Me.Conexao
            Comando.CommandType = CommandType.StoredProcedure
            Comando.CommandText = "SP_MM_INC_TB_LOG_OPER_ROBO"

            For Each papel In vListIdentPapeis
                If papel.CD_TITULO <> "" Then
                    vNu_Max_Seq += 1

                    With Comando.Parameters
                        .Clear()
                        .AddWithValue("@CDT_MOVIMENTO", Date.Now)
                        .AddWithValue("@CD_COTACAO", vIdentCot.CD_COTACAO)
                        .AddWithValue("@NU_SEQ", vNu_Max_Seq)
                        .AddWithValue("@CD_OPERACAO_PROD", vIdentCot.CD_OPERACAO_PROD)
                        .AddWithValue("@CDT_INICIO", vIdentCot.DT_INICIO)
                        .AddWithValue("@CDT_FIM", vIdentCot.DT_FIM)
                        .AddWithValue("@ENU_PRAZO_DU", vIdentCot.NU_PRAZO_DU)
                        .AddWithValue("@VL_TAXA", vIdentCot.VL_TAXA)
                        .AddWithValue("@VL_OPERACAO", vIdentCot.VL_OPERACAO)
                        .AddWithValue("@CD_INDEXADOR", vIdentCot.CD_INDEXADOR)
                        .AddWithValue("@CD_OPERADOR_COTACAO", vIdentCot.CD_OPERADOR_COTACAO)
                        .AddWithValue("@CD_CAIXA", papel.TipoEstoque)
                        .AddWithValue("@CD_PAPEL", papel.Cd_Papel)
                        .AddWithValue("@CD_TITULO", papel.CD_TITULO)
                        .AddWithValue("@CD_BACEN", papel.CD_BACEN)
                        .AddWithValue("@OT_VENCIMENTO", papel.DT_Vencimento)
                        .AddWithValue("@DT_VENCIMENTO_OPERACAO", papel.DT_Vencimento_Operacao)
                        .AddWithValue("@CD_ETQ", papel.Cd_ETQ)
                        .AddWithValue("@VL_PU", IIf(UCase(papel.TipoPrecoUnitario) = "PUMERCADO", papel.Vl_Pu_Mercado, papel.Vl_Pu_550))
                        .AddWithValue("@VL_PU_550", papel.Vl_Pu_Volta)
                        .AddWithValue("@NU_QUANTIDADE_TITULO", papel.QTDE_Titulo)
                        .AddWithValue("@NU_QUANTIDADE_UTILIZADA", papel.QTDE_Utilizada)
                        .AddWithValue("@NU_QUANTIDADE_DISPONIVEL", papel.Disponivel)
                        .AddWithValue("@DS_ERRO", "")
                    End With

                    Dim result As Integer = Comando.ExecuteNonQuery()

                    If result = 0 Then
                        Return False
                    End If
                End If
            Next
        End Using

        Return True

    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao gravar o log de operação do robô: " & ex.Message)
        Return False
    Finally
        Me.FecharConexao()
    End Try
End Function






'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################








Public Sub GravaLogOpersRobo (ByRef vpCD_COTACAO As String,
                                ByRef vpNu_Seq As Integer,
                                ByRef vpCD_OPERACAO_PROD As String, 
                                ByRef vpDT_INICIO As Date, 
                                ByRef VpDT_FIM As Date,
                                ByRef vpNU_PRAZO_DU As Integer,
                                ByRef vpVL_TAXA As Double,
                                ByRef vpVL_OPERACAO As Double,
                                ByRef vpCD_INDEXADOR As String,
                                ByRef vpCD_OPERADOR_COTACAO As String, 
                                ByRef vpCD_CAIXA As String,
                                ByRef vpCD_PAPEL As String,
                                ByRef vpCD_TITULO As String,
                                ByRef vpCD_BACEN As String,
                                ByRef vpDT_VENCIMENTO As String,
                                ByRef vpDT_VENCIMENTO_OPERACAO As String,
                                ByRef vpCD_ETQ As String,
                                ByRef vpVL_PU As String,
                                ByRef vpVL_PU_550 As String,
                                ByRef vpNU_QUANTIDADE TITULO As String, 
                                ByRef vpNU_QUANTIDADE UTILIZADA As String, 
                                ByRef vpNU_QUANTIDADE DISPONIVEL As String, 
                                ByRef vpDs_Erro As String)

        Dim Comando As New SqlCommand

        Try
            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "SP_MM_INC_TB_LOG_OPER_ROBO"
                
                With .Parameters
                    .Clear()
                    .AddWithValue("@DT_MOVIMENTO", Date.Now) 
                    .AddWithValue("@CD_COTACAO", vpCD_COTACAO) 
                    .AddWithValue("@NU_SEQ", vpNu_Seq)
                    .AddWithValue("@CD_OPERACAO_PROD", vpCD_OPERACAO_PROD) 
                    .AddWithValue("@DT_INICIO", vpDT_INICIO)
                    .AddWithValue("@DT_FIM", vpDT_FIM)
                    .AddWithValue("@NU_PRAZO_DU", vpNU_PRAZO_DU)
                    .AddWithValue("@VL_TAXA", vpVL_TAXA)
                    .AddWithValue("@VL_OPERACAO", vpVL_OPERACAO)
                    .AddWithValue("@CD_INDEXADOR", vpCD_INDEXADOR)
                    .AddWithValue("@CD_OPERADOR_COTACAO", vpCD_OPERADOR_COTACAO)
                    .AddWithValue("@CD_CAIXA", IIF(vpCD_CAIXA = "", DBNull.Value, vpCD_CAIXA))
                    .AddWithValue("@CD_PAPEL", IIf(vpCD_PAPEL = "", DBNull.Value, vpCD_PAPEL))
                    .AddWithValue("@CD_TITULO", IIf(vpCD_TITULO = "", DBNull.Value, vpCD_TITULO))
                    .AddWithValue("@CD_BACEN", IIF(vpCD_BACEN = "", DBNull.Value, vpCD_BACEN))
                    .AddWithValue("@DT_VENCIMENTO", IIF(vpDT_VENCIMENTO = "", DBNull.Value, vpDT_VENCIMENTO))
                    .AddWithValue("@DT_VENCIMENTO_OPERACAO", IIF(vpDT_VENCIMENTO_OPERACAO, DBNull.Value, vpDT_VENCIMENTO_OPERACAO)) 
                    .AddWithValue("@CD_ETQ", IIF(vpCD_ETQ = "", DBNull.Value, vpCD_ETQ))
                    .AddWithValue("@VL_PU", IIF(vpVL_PU = "", DBNull.Value, vpVL_PU))
                    .AddwithValue("@VL_PU_550", IIf(vpVL_PU_550 = "", DBNull.Value, vpVL_PU_550))
                    .AddWithValue("@NU_QUANTIDADE_TITULO", IIF(vpNU_QUANTIDADE_TITULO = "", DBNull.Value, vpNU_QUANTIDADE_TITULO))
                    .AddWithValue("@NU_QUANTIDADE_UTILIZADA", IIF(vpNU_QUANTIDADE_UTILIZADA = "", DBNull.Value, vpNU_QUANTIDADE_UTILIZADA)) 
                    .AddWithValue("@NU_QUANTIDADE_DISPONIVEL", IIF(vpNU_QUANTIDADEP_DISPONIVEL = "", DBNull.Value, vpNU_QUANTIDADE_DISPONIVEL)) 
                    .AddWithValue("@DS_ERRO", vpDs_Erro)

            End With

            .ExecuteNonQuery()
        End With

    Finally
        Me.FecharConexao()
    End Try

End Sub



Public Sub GravaLogOpersRobo(ByVal vpCD_COTACAO As String,
                              ByVal vpNu_Seq As Integer,
                              ByVal vpCD_OPERACAO_PROD As String,
                              ByVal vpDT_INICIO As Date,
                              ByVal vpDT_FIM As Date,
                              ByVal vpNU_PRAZO_DU As Integer,
                              ByVal vpVL_TAXA As Double,
                              ByVal vpVL_OPERACAO As Double,
                              ByVal vpCD_INDEXADOR As String,
                              ByVal vpCD_OPERADOR_COTACAO As String,
                              ByVal vpCD_CAIXA As String,
                              ByVal vpCD_PAPEL As String,
                              ByVal vpCD_TITULO As String,
                              ByVal vpCD_BACEN As String,
                              ByVal vpDT_VENCIMENTO As String,
                              ByVal vpDT_VENCIMENTO_OPERACAO As String,
                              ByVal vpCD_ETQ As String,
                              ByVal vpVL_PU As String,
                              ByVal vpVL_PU_550 As String,
                              ByVal vpNU_QUANTIDADE_TITULO As String,
                              ByVal vpNU_QUANTIDADE_UTILIZADA As String,
                              ByVal vpNU_QUANTIDADE_DISPONIVEL As String,
                              ByVal vpDs_Erro As String)

    Dim Comando As New SqlCommand

    Try
        Using Comando
            Comando.Connection = Me.Conexao
            Comando.CommandType = CommandType.StoredProcedure
            Comando.CommandText = "SP_MM_INC_TB_LOG_OPER_ROBO"

            With Comando.Parameters
                .Clear()
                .AddWithValue("@DT_MOVIMENTO", Date.Now)
                .AddWithValue("@CD_COTACAO", vpCD_COTACAO)
                .AddWithValue("@NU_SEQ", vpNu_Seq)
                .AddWithValue("@CD_OPERACAO_PROD", vpCD_OPERACAO_PROD)
                .AddWithValue("@DT_INICIO", vpDT_INICIO)
                .AddWithValue("@DT_FIM", vpDT_FIM)
                .AddWithValue("@NU_PRAZO_DU", vpNU_PRAZO_DU)
                .AddWithValue("@VL_TAXA", vpVL_TAXA)
                .AddWithValue("@VL_OPERACAO", vpVL_OPERACAO)
                .AddWithValue("@CD_INDEXADOR", vpCD_INDEXADOR)
                .AddWithValue("@CD_OPERADOR_COTACAO", vpCD_OPERADOR_COTACAO)
                .AddWithValue("@CD_CAIXA", If(String.IsNullOrEmpty(vpCD_CAIXA), DBNull.Value, vpCD_CAIXA))
                .AddWithValue("@CD_PAPEL", If(String.IsNullOrEmpty(vpCD_PAPEL), DBNull.Value, vpCD_PAPEL))
                .AddWithValue("@CD_TITULO", If(String.IsNullOrEmpty(vpCD_TITULO), DBNull.Value, vpCD_TITULO))
                .AddWithValue("@CD_BACEN", If(String.IsNullOrEmpty(vpCD_BACEN), DBNull.Value, vpCD_BACEN))
                .AddWithValue("@DT_VENCIMENTO", If(String.IsNullOrEmpty(vpDT_VENCIMENTO), DBNull.Value, vpDT_VENCIMENTO))
                .AddWithValue("@DT_VENCIMENTO_OPERACAO", If(String.IsNullOrEmpty(vpDT_VENCIMENTO_OPERACAO), DBNull.Value, vpDT_VENCIMENTO_OPERACAO))
                .AddWithValue("@CD_ETQ", If(String.IsNullOrEmpty(vpCD_ETQ), DBNull.Value, vpCD_ETQ))
                .AddWithValue("@VL_PU", If(String.IsNullOrEmpty(vpVL_PU), DBNull.Value, vpVL_PU))
                .AddWithValue("@VL_PU_550", If(String.IsNullOrEmpty(vpVL_PU_550), DBNull.Value, vpVL_PU_550))
                .AddWithValue("@NU_QUANTIDADE_TITULO", If(String.IsNullOrEmpty(vpNU_QUANTIDADE_TITULO), DBNull.Value, vpNU_QUANTIDADE_TITULO))
                .AddWithValue("@NU_QUANTIDADE_UTILIZADA", If(String.IsNullOrEmpty(vpNU_QUANTIDADE_UTILIZADA), DBNull.Value, vpNU_QUANTIDADE_UTILIZADA))
                .AddWithValue("@NU_QUANTIDADE_DISPONIVEL", If(String.IsNullOrEmpty(vpNU_QUANTIDADE_DISPONIVEL), DBNull.Value, vpNU_QUANTIDADE_DISPONIVEL))
                .AddWithValue("@DS_ERRO", vpDs_Erro)
            End With

            Comando.ExecuteNonQuery()
        End Using

    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao gravar o log de operação do robô: " & ex.Message)
    Finally
        Me.FecharConexao()
    End Try

End Sub





'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################







Public Function GravaComplementoTitpu(ByVal vIdentCot As Contrato.IdentificarCotacao, 
                                        ByVal vNu_Max_Seq As Integer,
                                        ByVal vListIdentPapeis As List (Of Contrato.PapeisUsadosParaLastro)) As Boolean

        Dim Comando As New SqlCommand
        Dim Retorna As String
        Dim y As Integer
        Dim contas As ContasLiquidacao



        Try

            With Comando
                .Connection = Me.Conexao
                .CommandType = CommandType.StoredProcedure 
                .CommandText = "SP_INSERE COMPL_TP"
            End With

            For y = 0 To vListIdentPapeis.Count - 1

                If vListIdentPapeis(y).CD_TITULO <> "" Then

                    With Comando
                    vNu_Max_Seq += 1

                    contas = RetornaContasBaseadoCdEtqPapel(vIdentCat, vListIdentPapeis(y))


                    With Parameters
                        .Clear()
                        .AddWithValue("@CD_OPERACAO", vIdentCot.CD_COTACAO) 
                        .AddWithValue("@NU_SEQ", vNu_Max_Seq) 
                        .AddWithValue("ONU_SEQ_PERNA", 0)
                        .AddWithValue("@CD_EMPRESA", vIdentCot.CD_EMPRESA) 
                        .AddWithValue("@CD_INDEXADOR_OPERACAO", "NO") 
                        .AddWithValue("@CD_INDEXADOR_ORIGEM", "NO") 
                        .AddWithValue("@CD_TIPO_LASTRO_OPERACAO", "SLN") 
                        .AddWithValue("@CD_TIPO_LASTRO_ORIGEM", "NO")
                        .AddWithValue("@NM_EMPRESA", vIdentCot. NM_EMPRESA) 
                        .AddWithValue("@CAMARA_CODIGO", 1)
                        .AddWithValue("@IC_ESTOQUE_TRIGGER", 1)
                        .AddWithValue("@IC_BOLETA", 1)
                        .AddWithValue("@CD_ETQ", vListIdentPapeis (y).cd_ETQ)
                        .AddWithValue("QIC_LIVRE_MOVIMENTACAO", vIdentCot. IC_LIVRE_MOVIMENTACAO) 
                        .AddWithValue("@CNPJ_empresa", vIdentCot.NU_ISPB_IF)
                        .AddWithValue("@cd_Sistema_HT", vIdentCot.CD_SISTEMA_HT)
                        .AddWithValue("@cd_Sistema_Origem", vIdentCot.CD_SISTEMA_ORIGEM)
                        .AddWithValue("@IC_CONTRAPARTE_BROKER", vIdentCot.IC_CONTRAPARTE_BROKER) 
                        .AddWithValue("@IC_ORIGEM", 1)
                        .AddWithValue("@Cta_Selic_Cedente", contas.ContaCedente)
                        .AddWithValue("@Cta_Selic_Liquid_Cedente", contas.ContaLiquidacaoCedente) 
                        .AddWithValue("@Cta_Selic_Cessionario", contas.ContaCessionario)
                        .AddWithValue("@Cta_Selic_Liquid Cessionario", contas.ContaLiquidacaoCessionario)

                    End With

                    Retorna.ExecuteScalar().ToString

                End With

                If Retorna >>"0" Then

                Return False

            End If
        Next

        Return True

    Finally
        Me.FecharConexao()
    End Try

End Function



Public Function GravaComplementoTitpu(ByVal vIdentCot As Contrato.IdentificarCotacao,
                                       ByVal vNu_Max_Seq As Integer,
                                       ByVal vListIdentPapeis As List(Of Contrato.PapeisUsadosParaLastro)) As Boolean

    Dim Comando As New SqlCommand
    Dim Retorna As String
    Dim contas As ContasLiquidacao

    Try
        Using Comando
            Comando.Connection = Me.Conexao
            Comando.CommandType = CommandType.StoredProcedure
            Comando.CommandText = "SP_INSERE_COMPL_TP"

            For Each papel In vListIdentPapeis
                If papel.CD_TITULO <> "" Then
                    vNu_Max_Seq += 1
                    contas = RetornaContasBaseadoCdEtqPapel(vIdentCot, papel)

                    With Comando.Parameters
                        .Clear()
                        .AddWithValue("@CD_OPERACAO", vIdentCot.CD_COTACAO)
                        .AddWithValue("@NU_SEQ", vNu_Max_Seq)
                        .AddWithValue("@NU_SEQ_PERNA", 0)
                        .AddWithValue("@CD_EMPRESA", vIdentCot.CD_EMPRESA)
                        .AddWithValue("@CD_INDEXADOR_OPERACAO", "NO")
                        .AddWithValue("@CD_INDEXADOR_ORIGEM", "NO")
                        .AddWithValue("@CD_TIPO_LASTRO_OPERACAO", "SLN")
                        .AddWithValue("@CD_TIPO_LASTRO_ORIGEM", "NO")
                        .AddWithValue("@NM_EMPRESA", vIdentCot.NM_EMPRESA)
                        .AddWithValue("@CAMARA_CODIGO", 1)
                        .AddWithValue("@IC_ESTOQUE_TRIGGER", 1)
                        .AddWithValue("@IC_BOLETA", 1)
                        .AddWithValue("@CD_ETQ", papel.Cd_ETQ)
                        .AddWithValue("@QIC_LIVRE_MOVIMENTACAO", vIdentCot.IC_LIVRE_MOVIMENTACAO)
                        .AddWithValue("@CNPJ_empresa", vIdentCot.NU_ISPB_IF)
                        .AddWithValue("@cd_Sistema_HT", vIdentCot.CD_SISTEMA_HT)
                        .AddWithValue("@cd_Sistema_Origem", vIdentCot.CD_SISTEMA_ORIGEM)
                        .AddWithValue("@IC_CONTRAPARTE_BROKER", vIdentCot.IC_CONTRAPARTE_BROKER)
                        .AddWithValue("@IC_ORIGEM", 1)
                        .AddWithValue("@Cta_Selic_Cedente", contas.ContaCedente)
                        .AddWithValue("@Cta_Selic_Liquid_Cedente", contas.ContaLiquidacaoCedente)
                        .AddWithValue("@Cta_Selic_Cessionario", contas.ContaCessionario)
                        .AddWithValue("@Cta_Selic_Liquid_Cessionario", contas.ContaLiquidacaoCessionario)
                    End With

                    Retorna = Comando.ExecuteScalar().ToString()

                    If Retorna <> "0" Then
                        Return False
                    End If
                End If
            Next

            Return True
        End Using

    Catch ex As Exception
        ' Lidar com exceção aqui
        Console.WriteLine("Ocorreu um erro ao gravar o complemento de título: " & ex.Message)
        Return False
    Finally
        Me.FecharConexao()
    End Try

End Function


 

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################




Public Function AlterarCotacaoMae (Byval vpCdCotacao As String, Byval vpStatus As Statusopvinculo) As Boolean 
    Dim Comando As New SqlCommand
    Dim Retorna As Integer

    Try
    With Comando
        Connection = Me.Conexao
        .CommandType = CommandType.StoredProcedure 
        .CommandText = "SP_MM_ALT_COTACAO_ORIGEM"
        With .Parameters
            .Clear()
            .AddWithValue("@cd_cotacao", vpCdCotacao) 
            .AddWithValue("@cd_status", vpStatus)
        End With

        Retorna CType(.ExecuteNonQuery(), Integer)

        If Retorna = 0 Then
            Return False
        Else
            Return True
        End If

    End With

    Finally
        Me.FecharConexao()
    End Try

End Function




Public Function AlterarCotacaoMae(ByVal vpCdCotacao As String, ByVal vpStatus As Statusopvinculo) As Boolean 
    Dim Comando As New SqlCommand

    Try
        With Comando
            .Connection = Me.Conexao
            .CommandType = CommandType.StoredProcedure 
            .CommandText = "SP_MM_ALT_COTACAO_ORIGEM"

            With .Parameters
                .Clear()
                .AddWithValue("@cd_cotacao", vpCdCotacao) 
                .AddWithValue("@cd_status", vpStatus)
            End With

            Return .ExecuteNonQuery() <> 0
        End With

    Catch ex As Exception
        ' Lida com a exceção aqui
         Console.WriteLine("Ocorreu um erro ao alterar a AlterarCotacaoMae: " & ex.Message)
        Return False
    End Try
End Function

 

'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################


Public Function VerificaEstoqueDisponivel(Byval vpDataMesa As Date,
                                            Byval vIdentCot As Contrato.IdentificarCotacao,
                                            Byval vListIdentPapeis As List(of Contrato.PapeisUsadosParaLastro)) As Boolean
    Dim Comando As New SqlCommand
    Dim Retorna As String
    Dim y As Integer
    Try
        
        For y = To vListIdentPapeis.Count - 1
            If vListIdentPapeis(y).CD_TITULO <> "" Then
                With Comando
                    With .Parameters 
                        .Clear()
                        
                        .AddithValue("@codCanara", vIdentCot.CD_CAMARA) 
                        .AddWithValue("@nm_mesa", vIdentCot.NPRESA) 
                        .AddWithValue("@codEmp", vIdentCot.CD_EMPRESA)

                        If vListIdentPapeis(y).TipoEstoque = "T" Then 
                            .AddWithValue("@codCai", "TERCEIROS")
                            .AddWithValue("edtVenorg", vListIdentPapeis (y).DT_Vencimento_Operacao.ToString("yyyyMMdd")) 
                        ElseIf vListIdentPapeis (y). TipoEstoque = "p" Then
                            .AddWithValue("@codCai", "PROPRIO")
                            .AddhrithValue("@dtVenorg", vListIdentPapeis (y).DT_Vencimento.ToString("yyyyMMdd"))
                        ElseIf vListIdentPapeis (y).TipoEstoque = "L" Then
                            .AddithValue("@codCai", "LA")
                            .AdithValue("@dtVenorg", vListIdentPapeis (y).DT_Vencimento. TeString("yyyyMMdd"))
                        End If

                            .AddithValue("@papel", viistidentPapais(y).Gil_Pape1)
                            .AddidithValue("@dtIni", videntCat.DT_INICIO.TeString("yyyyMMdd"))
                            .AddicithValue("@dtVen", videntCat.DT_FIN.TeString("yyyyMMdd"))

                            'Na verificacao do estoque disponivel somar a quantidade reservada
                            .AddWithValue("@qtd", vListIdentPapeis(y).QTDE_Utilizada + vListIdentPapeis(y).QTDE_RESERVADA) 
                            .AddWithValue("@dtJor", vpDataMesa.ToString("yyyyMMdd"))
                            '--AddWithValue("@dtLiq", )
                            .AddWithValue("@codEtq", vListIdentPapeis(y).cd_ETQ)
                            
                            Dim stInicio As Date = vIdentCot.DT_INICIO
                            Dim stTermino As Date = vIdentCot.DT_FIM

                            If vIdentCot.CD_INDEXADOR = "PRE" And vIdentCot.NU_PRAZO_DU = 1 Then

                                Dim sDtVencto As Date = CDate(vListIdentPapeis(y).DT_Vencimento)
                                Dim sprxVenctoPapel As Date

                                If Not Datas.DiaUtil(sDtVencto) Then '--False = não é dia útil, True = é dia útil
                                    sDtVencto = Datas. ObterProximoDiaUtil(sDtVencto)
                                    .AddWithValue("@DiaUtil", 1)
                                    .AddWithValue("@PrxDtVenOrg", sDtVencto.ToString("yyyyMMdd"))
                                Else
                                
                                    sPrxVenctoPapel = sDtVencto
                                    .AddWithValue("@DiaUtil", 0)
                                End If
                            Else
                                .AddWithValue("@DiaUtil", 0)
                            End If

                        End With

                        .Connection = Me.Conexao
                        .CommandType = CommandType.StoredProcedure

                        If vListIdentPapeis (y). TipoEstoque = "L" Then 
                            .CommandText = "SP_MM_VERIFICA_QTD_ESTOQUE_LM"
                        Else
                            .CommandText = "SP_MM_VERIFICA_QTD_ESTOQUE_PROPRIO"
                        End If

                        Retorna = .ExecuteScalar().ToString

                    End With
                    
                    If Retorna <> "0" Then
                        Return False
                    End If
            End If
        Next

        Return True

    Finally
        Me.FecharConexao()
    End Try

End Function





Public Function VerificaEstoqueDisponivel(ByVal vpDataMesa As Date,
                                           ByVal vIdentCot As Contrato.IdentificarCotacao,
                                           ByVal vListIdentPapeis As List(Of Contrato.PapeisUsadosParaLastro)) As Boolean
    Try
        For Each papeis As Contrato.PapeisUsadosParaLastro In vListIdentPapeis
            If papeis.CD_TITULO <> "" Then
                Using Comando As New SqlCommand
                    With Comando.Parameters
                       .Clear()
                    .AddWithValue("@codCanara", vIdentCot.CD_CAMARA)
                    .AddWithValue("@nm_mesa", vIdentCot.NPRESA)
                    .AddWithValue("@codEmp", vIdentCot.CD_EMPRESA)

                    If papeis.TipoEstoque = "T" Then
                        .AddWithValue("@codCai", "TERCEIROS")
                        .AddWithValue("@dtVenorg", papeis.DT_Vencimento_Operacao.ToString("yyyyMMdd"))
                    ElseIf papeis.TipoEstoque = "p" Then
                        .AddWithValue("@codCai", "PROPRIO")
                        .AddWithValue("@dtVenorg", papeis.DT_Vencimento.ToString("yyyyMMdd"))
                    ElseIf papeis.TipoEstoque = "L" Then
                        .AddWithValue("@codCai", "LA")
                        .AddWithValue("@dtVenorg", papeis.DT_Vencimento.ToString("yyyyMMdd"))
                    End If

                    .AddWithValue("@papel", papeis.Cd_Papel)
                    .AddWithValue("@dtIni", vIdentCot.DT_INICIO.ToString("yyyyMMdd"))
                    .AddWithValue("@dtVen", vIdentCot.DT_FIN.ToString("yyyyMMdd"))
                    .AddWithValue("@qtd", papeis.QTDE_Utilizada + papeis.QTDE_RESERVADA)
                    .AddWithValue("@dtJor", vpDataMesa.ToString("yyyyMMdd"))

                    Dim stInicio As Date = vIdentCot.DT_INICIO
                    Dim stTermino As Date = vIdentCot.DT_FIM

                    If vIdentCot.CD_INDEXADOR = "PRE" And vIdentCot.NU_PRAZO_DU = 1 Then
                        Dim sDtVencto As Date = CDate(papeis.DT_Vencimento)
                        Dim sprxVenctoPapel As Date

                        If Not Datas.DiaUtil(sDtVencto) Then
                            sDtVencto = Datas.ObterProximoDiaUtil(sDtVencto)
                            .AddWithValue("@DiaUtil", 1)
                            .AddWithValue("@PrxDtVenOrg", sDtVencto.ToString("yyyyMMdd"))
                        Else
                            sprxVenctoPapel = sDtVencto
                            .AddWithValue("@DiaUtil", 0)
                        End If
                    Else
                        .AddWithValue("@DiaUtil", 0)
                    End If
                End With

                    With Comando
                        .Connection = Me.Conexao
                        .CommandType = CommandType.StoredProcedure

                        If papeis.TipoEstoque = "L" Then
                            .CommandText = "SP_MM_VERIFICA_QTD_ESTOQUE_LM"
                        Else
                            .CommandText = "SP_MM_VERIFICA_QTD_ESTOQUE_PROPRIO"
                        End If

                        Return .ExecuteScalar().ToString() = "0"
                    End With
                End Using
            End If
        Next

        ' Se nenhum estoque indisponível for encontrado, retorna verdadeiro
        Return True

    Catch ex As Exception
        ' Tratar exceção aqui
        Return False

    Finally
        Me.FecharConexao()
    End Try
End Function




'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################

 



Public Function LocalizaCodSerie(ByVal vpCd_Papel As Integer) As Integer
    Dim vComando As New SqlCommand
    Dim retorna As Integer
    Dim SSQL As String

    SSQL = "Select Isnull(COD_SRIE, 0) From TB PAPEL (NOLOCK) Where CD_PAPEL=" & vpcd_Papel & "" 

    Try
        With vComando
            .Connection = Me.Conexao
            .CommandType = CommandType.Text
            .CommandText = SSQL
        End With

        retorna = CType(vComando.ExecuteScalar(), Integer)

        Return retorna

    Finally
        Me.FecharConexao()
    End Try
End Function




Public Function LocalizaCodSerie(ByVal vpCd_Papel As Integer) As Integer
    Dim retorna As Integer = 0
    Dim SSQL As String = "SELECT ISNULL(COD_SRIE, 0) FROM TB_PAPEL WITH (NOLOCK) WHERE CD_PAPEL = @CdPapel"

    Try
        Using conexao As New SqlConnection(Me.Conexao)
            conexao.Open()

            Using comando As New SqlCommand(SSQL, conexao)
                comando.Parameters.AddWithValue("@CdPapel", vpCd_Papel)

                Dim reader As SqlDataReader = comando.ExecuteReader()
                If reader.Read() Then
                    retorna = Convert.ToInt32(reader(0))
                End If
            End Using
        End Using

    Catch ex As Exception
        ' Lidar com a exceção aqui, se necessário
        retorna = 0

    Finally
        Me.FecharConexao()
    End Try

    Return retorna
End Function




'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################



Public Function TipoPrecoUnitario(ByVal vpCDCotacao As String) As String
    Dim vComando As New SqlCommand
    Dim retorna As String
    Dim sSQL As String

    SSQL = "select T830.DES_TIPO_PREC_UNIT from TBMM826_CMPL_OPER_PREC_UNIT TB26 INNER JOIN TBMM830_TIPO_PREC_UNIT TB30 ON TB26.COD_TIPO_PREC_UNIT = TB30.COD_TIPO_PREC_UNIT WHERE TB26.CD_OPERACAO='" & vpCDCotacao & "'"
    
    Try
    
        With vComando
            .Connection = Me.Conexao
            .CommandType = CommandType.Text
            .CommandText = SSQL
        End With

        retorna CType(vComando.ExecuteScalar(), String)

        Return UCase(Trim(retorna))

    Finally
        Me.FecharConexao()
    End Try
End Function




Public Function TipoPrecoUnitario(ByVal vpCDCotacao As String) As String
    Dim retorna As String = ""
    Dim sSQL As String = "SELECT T830.DES_TIPO_PREC_UNIT FROM TBMM826_CMPL_OPER_PREC_UNIT TB26 INNER JOIN TBMM830_TIPO_PREC_UNIT TB30 ON TB26.COD_TIPO_PREC_UNIT = TB30.COD_TIPO_PREC_UNIT WHERE TB26.CD_OPERACAO = @CdCotacao"

    Try
        Using conexao As New SqlConnection(Me.Conexao)
            conexao.Open()

            Using comando As New SqlCommand(sSQL, conexao)
                comando.Parameters.AddWithValue("@CdCotacao", vpCDCotacao)

                Dim reader As SqlDataReader = comando.ExecuteReader()
                If reader.Read() Then
                    retorna = Convert.ToString(reader("DES_TIPO_PREC_UNIT")).Trim()
                End If
            End Using
        End Using

    Catch ex As Exception
        ' Lidar com a exceção aqui, se necessário

    Finally
        Me.FecharConexao()
    End Try

    Return retorna.ToUpper()
End Function





'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################
'######################################################################################################################################################################











