Namespace Servico  

 
    Public Class Lastro  
        Inherits ServicoBase  


        Public Const vCdEmpresa As String $=" 010 "$  
        Public Const vilesa As String = "MMERC"  
        Public Const TipoDeBoletador As String = "VL"  
        Public Const vUsuarioControln As String = "CTRLMVA"  
        Public Const vCdTaxa As String = "OPCONP"  
        
        
        Public Papeisutilizados As New List(of Contrato.PapeisUsadosParaLastro)  

        Dim vAcessoDados As Mew AcessoDados.Configuracao  

        Public Function ExecutarLastro(Byval Datahesa As Date) As Boolean  


            Dim ResultadoParanetrizacaopapetslastio As List(of Contrato. ParanetrizacaoPapeislastro)  
            Dim ResultadoSelecionarCotacoes As List(of Contrato.OperacoesCompromissadas)  
            Dim ResultadoDadostbTaxaConpronissada As Contrato.DadosTbTaxaConpromissada  

            Dim ResultadotaxaConprts As Nen List(of Contrato.Dadostbraxacompronissada)  

            Dim vsValorcalc As Double  
            Dim vsMtenorPu As Double  
            Dim blOperacao As Boolean

            'Verifica o cadastranento da taxa de operaçōes compromissadas - Data da Mesa  
            ResultadoDadosTbTaxaCompromissada = vAcessoDados.VerificaTaxaCompromissada(vCdEmpresa, DataMesa, vCdTaxa)  

            ResultadoTaxaCompris.Add(ResultadoDadosTbTaxaCompromissada)  

            Dim AntDthesa As Date = Itau.MM.Franemork.GerenciadorDatas.Datas.ObterUltinoDiautil(DataMesa.AddDays(-1))  

            'Verifica o cadastramento da taxa de operaçōes conpronissadas - Data da Mesa  
            ResultadoDadosTbTaxaCompromissada = vAcessoDados.VerificaTaxaCompromissada(vCdEmpresa, AntDtMesa, vCdraxa)  

            ResultadoTaxaCompris.Add(ResultadoDadosTbTaxaCompromissada)  

            'Seleciona Operaçōes Conpronissadas - 0540D - Pre - Hercado - (SIST_HT, SIST_LM, SIST_VW)  
            ResultadoSelecionarCotacoes = vAcessoDados.SelectonarCotacoes(vedanpresta, "", 0 , Contrato.TipoOrigemOperacao.Robolastio)  
            If ResultadoSelecionarCotacoes.Count = 0 Then 
                Return False  
            End If  

            Dim sMsg As String  
            Dim Achoutaxa As Integer =  0


            For Each vitemOp As Contrato.OperacoesCompromissadas In ResultadoSelecionarCotacoes  
                Logger.GerarLog("Executando o lastro: " & vItemOp.cd_Cotacao)  

                'Paranetrizaçāo default para estoque de terceiros  
                ResultadoParametrizacaoPapeisLastro = vAcessoDados.BuscaParametrizacaoLastro(vCdEmpresa, vItenOp.dt_inicio, vItenOp.CD_OPERADOR_COTACAO)  

                For Each vItemTaxaCompromis As Contrato.DadosTbTaxaCompromissada In ResultadoTaxaCompris  
                    If ResultadoParametrizacaoPapeisLastro.Count = 0 Then  
                        'Näo ha paranetrizacao de papeis para 0 dia/sist_origen  
                        AchouTaxa = 4  
                        Exit For  
                    Else  
                        If vItenOp.VL_OPERACAO > ResultadoParanetrizacaoPapeisLastro.Iten(0).VL_FINAMC Then  
                            'Se o valor da operaçtio for naior que o parametrizado  
                            Achoutaxa = 3  
                            Exit For  
                        Else  
                            If vItemOp,dt_inicio.ToString("yyyyHidd") = vItenTaxaCompronis.DT_REF Then  
                                AchouTaxa = 2  
                                Exit For  
                            End If  
                        End If  
                    End If 
                Next


                'Näo ha paranetrizacao de papeis para o dia/sist_origen  
                If AchouTaxa = 4 Then  
                vAcessoDados.GravaLogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vItemOp.dt_inicio, vItemOp.DT_FIM,
                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO,
                                                "","","","","",
                                                "","","","","",
                                                "","",
                                                "Verificar Paranetrizacao de Papeis para o dia" & vItemOp.dt_inicio & "_" & vItemOp.CD_OPERADOR_COTACAO )
                
                End If
                
                'Se o valor da operaçâo for maior que o paranetrizado  
                If Achoutaxa = 3 Then  
                vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vitemOp.dt_inicio, vItemOp.DT_FIM, 
                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO,
                                                "","","","","",
                                                "","","","","",
                                                "","",
                                                "Valor da Operacao > Parametrizado Valor Operacao = " & vItemOp.VL_OPERACAO & " - Valor Prametrizacao = VVVV")
                
                End If  
                
                'Se a taxa da compronissada nâo for localizada para o dia de inicio da operaçäo  
                If Achoutaxa = 0 Then  
                vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vitemOp.dt_inicio, vitemOp.DT_FIM, 
                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO,
                                                "","","","","",
                                                "","","","","",
                                                "","",
                                                "Tara de Compronissada nâo cadastrada para o dia" & vItemOp.dt_inicio)
                
                End If  


                'Se a taxa da compromissada for diferente da taxa da operacao desprezar  
                If AchouTaxa = 1 Then  
                vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vitemOp.dt_inicio, vitemOp.DT_FIM, 
                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO,
                                                "","","","","",
                                                "","","","","",
                                                "","",
                                                "Valor da taxa " & vItemOp.vl_taxa & " fora do delta permitido ")
                End If

                If Achoutaxa = 2 Then  
                    with vItemOp  
                        'Bloquear cliente  
                        sMsg = ""  
                        If Not .MN_CLTEMrE = "BACEMTES" Then  
                            If Not vAcessoDados.sBloquearDesbloquearcltente(vUsuarioControlM,
                                                                            cd_cliente,  
                                                                            NM_CLIENTE,
                                                                            vpTipoAcao:=Contrato.sBloquearDesbloquearcltente.bdcBloquear,
                                                                            vpMsg:=sMsg) Then  
                            '-sing = "Erro ao Bloquear o cliente"  
                            End If  
                        End If

                        If sMsg = "" Then  
                            vsvalorCalc = 0 
                            vsMenorPu = 0  
                            PapeisUtilizados.Clear()  
                            For Each vItem As Contrato.ParanetrizacaoPapeisLastro In ResultadoParametrizacaoPapeisLastro  

                                If LastreamentoAutomatico(vItemOp, vItem, vsValorCalc, vsMenorPu) Then
                                    If Abs(vItemOp.VL_OPERACAO - vsValorCalc) < vsMenorPu And vsValorCalc > 0 Then Exit For  
                                End If

                            Next 

                            If Abs(vItemOp.VL_OPERACAO - vsValorCalc) < vsMenorPu And vsValorCalc > 0 Then 

                                blOperacao = True  
                                For Each pMercado As Contrato.PapeisUsadosParalastro In PapeisUtilizados  
                                    If pMercado.TipoPU_Operacao = "PUMERCADO" And pMercado.TipoPrecoUnitario = "PU550" Then  
                                        blOperacao = False  
                                        Exit For  
                                    End If 
                                Next
                              
                                If blOperacao = True Then 
                                    Grava_Complemento_Cotacao(DataMesa, vItemOp.cd_Cotacao)  
                                Else
                                    Dim vIdentCot As Contrato.IdentificarCotacao  
                                    vIdentCot = vAcessoDados.IdentificarCotacao(vItemOp.cd_Cotacao)  
                                    PapeisUtilizados(0).TipoPrecoUnitario = "PU550"  
                                    If Not vAcessoDados.GravaCMPL_OPER_PREC_UNIT(vIdentCot, 0, PapeisUtilizados) Then
                                        vAcessoDados.GravalogOpersRobo( vItemOp.cd_Cotacao, 0, "0540D",  
                                                                vItemOp.dt_inicio, vItemOp.DT_FIM,  
                                                                vItemOp.nu_prazo_du, vItemOp.vl_taxa, vItemOp.VL_OPERACAO,  
                                                                vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO, 
                                                                "","","","","",
                                                                "","","","","",
                                                                "","",
                                                                "Erro ao Gravar Tabela TBrn026_CrPL_OPER_PREC_UNIT na mudança TipoPu. ")


                                        Exit Function  
                                    End If
                                    vAcessoDados.GravalogOpersRobo( vItemOp.cd_Cotacao, 0, "0540D", vItemOp.dt_inicio, vItemOp.DT_FIM,  
                                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO, 
                                                                "","","","","",
                                                                "","","","","",
                                                                "","",
                                                                "TipopU da operapio diverge dos lastros. Financianento nfio confinado !!!")
                                End If  
                
                            Else  
                                vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vItemOp.dt_inicio, vItemOp.DT_FIM,  
                                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO, 
                                                                "","","","","",
                                                                "","","","","",
                                                                "","",
                                                                "Nấo há quantidade suficiente disponivel para lastrear a operação. Financianento não confirmado  !!!")
                    
                            End If


                            'Desbloquear cliente  
                            sMsg = ""  
                            If Not .NM_CLIENTE.ToString = "BACENTES" Then  
                                If Not vAcessoDados.sBloquearDesbloquearCliente(vUsuarioControlM,  
                                                                                .cd_cliente,  
                                                                                .NM_CLIENTE,  
                                                                                vpTipoAcao:=Contrato.BloquearDesbloquearCliente.bdcDesBloquear,  
                                                                                vpMsg:=sMsg) Then  
                                sMsg = "Erro ao Desbloquear o cliente"   
                                End If  
                            End If
                        Else
                            vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "0540D", vItemOp.dt_inicio, vItemOp.DT_FIM,  
                                                                1, vItemOp.vl_taxa, vItemOp.VL_OPERACAO, vItemOp.Cd_Indexador, vItemOp.CD_OPERADOR_COTACAO, 
                                                                "","","","","",
                                                                "","","","","",
                                                                "","",
                                                                sMsg)
                        End If
                    End with
                End If
            Next
        End Function


    Public Function LastreamentoAutomatico(ByRef vItemOp As Contrato.OperacoesCompromissadas,  
                                            ByRef vItemPapeis As Contrato. ParametrizacaoPapeisLastro,  
                                            ByRef vpValorCalc As Double,  
                                            ByRef vpMenorPU As Double  
                                            ) As Boolean  
        Dim vsQtdeBase As Integer  
        Dim vsQtde As Integer  
        Dim vsValorOrig As Double  
        Dim vsValorLastro As Double  

        Dim vsValorOperacao As Double  

        vsValorOrig = vItemOp.VL_OPERACAO  
        vsValorOperacao = vItemOp.VL_OPERACAO  

        'Area para guardar os titulos para determinado estoque  
        Dim ResultadoPapeis As List(of Contrato.PapeisUsadosParaLastro)  

        'busca a paranetrizacao para saber se as operaçōes serāo lastreadas pelo PU550 ou PUHer qado  
        'valores esperados : PUnERCADO e PU550  
        Dim TipoPrecoUnitario As String  
        TipoPrecoUnitario = vicessobados.TipoPrecoUnitario(vItemOp.cd_cotacao)  

        Logger.GerarLog("LastreanentoAutomatico")


        With vItemPapeis  
            ResultadoPapeis = vAcessoDados.TitulosParaLastro(.ID_ESTOQUE, vMesa, vCdEmpresa, vItemOp.dt_inicio, vItemOp.DT_FIM, , , Contrato.FormaLastro.flRobo, Contrato.TipoOrigemOperacao.Integrador, vItemOp.CD_OPERADOR_COTACAO) 
            
            If ResultadoPapeis.Count <= 0 Then  
                vAcessoDados.GravalogOpersRobo(vItemOp.cd_Cotacao, 0, "05400", vItemOp.dt_inicio, vItemOp.DT_FIM,  
                                                .ID_ESTOQUE , "", "", "", "",  
                                                "", "", "", "", "",
                                                "", "",
                                                "Não há papéis para lastrear. ")
                Return False  
            End If  
        End with

        'Catía - Pu Titulos Publicos  
        If UCase(tipoPrecoUnitario) = "PUMERCADO" Then  
            'ResultadoPapeisMercado. Clear()  
            'ResultadoPapeis Mercado. AddRange (ResultadoPapeis)  
            Logger.GerarLog("ConsultaServicoIM")  
            For Each pMercado As Contrato.PapeisUsadosParaLastro In ResultadoPapeis  
            ConsultaServicoIM(pMercado, vItemOp.dt_inicio, vItemOp)  
            Next
        End If


        Dim blachei As Boolean = False 
        Dim ValorTitulo As Double = 0 
        Dim intIndex As Integer = 0
        Dim DadosPU550 As New Contrato.DadosPU550 
        Dim rtvl Pu_Volta As Double
        Dim rtvl Pu_550 As Double


        For Each p As Contrato.PapeisUsadosParaLastro In ResultadoPapeis
            p.TipoPU_Operacao UCase(tipoPrecoUnitario)
            If p.VI_Pu_550 = 0 Then
                DadosPUSSO = vAcessoDados.ConsultaPUS550(p.Cd_Papel)
                p.VL_Pu_Volta = DadosPUS50.VL_PU_RET
                p.VL_Pu_550 = Trunca(8, DadosPU550.VL_PU_RET / (((CDbl(vItemOp.vl_taxa) / 100 + 1) ^ (Val(vItemOp.nu_prazo_du) / 252))))
                rtvl_Pu_Volta = 0
                rtvl_Pu_550 = 0
                ConfirmaValorPUna_Vespera_Vencto_Papel_(p.VL_P_Volta, p.VL_Pu_550, vItemOp.vl_taxa, vItemOp.nu_prazo_du, rtvl_Pu_Volta, rtvl_Pu_550) 
                p.VL_Pu_Volta = rtvl_Pu_Volta
                p.VI_Pu_550 = rtvl_Pu_550
                p.Valor = Int(p.Vl_Pu_550 * p.Disponivel)
                If UCase(tipoPrecoUnitario) = "PUMERCADO" Then
                    p.VL_Pu_Mercado = rtvl_Pu_550
                    p.TipoPrecoUnitario = "PUMERCADO"
                Else
                    p.TipoPrecoUnitario = "PUS50"
                End If    
            Else
                If UCase(tipoPrecoUnitario) = "PUMERCADO" And p.VL_PuMercado > 0 Then
                    p.VL_Pu_Volta = Trunca(8, p.VL_Pu_Mercado * (((CDbl(vItemOp.vl_taxa) / 100+ 1) ^ (Val(vItemOp.nu_praze_du) / 252))))
                    p.TipoPrecoUnitario = "PUMERCADO"
                Else
                    p.VL_Pu_Volta = Trunca(8, p.VL_Pu_550 * (((CDbl(vItemOp.vl_taxa) / 100+ 1) ^ (Val(vItemOp.nu_praze_du) / 252)))) 
                    p.TipoPrecoUnitario = "PUSSO"
                End If
            
            End If
        Next





        While Abs(vsValorOrig - vpValorCalc) > vpMenorPU
            blachei = False 
            ValorTitulo = 0
            intIndex = 0
        
            'Procurar un titulo onde o valor disponivel deve ser ao valor da operacao
            CalculaQtde(ResultadoPapeis, Abs(vsValorOrig - vpValorCalc), 1, blachei, ValorTitulo, intIndex)

            'Se não achei, voltar a procurar o maior valor disponivel
            If blachei = False Then
                blachei = False 
                ValorTitulo = 0
                intIndex = 0
                CalculaQtde(ResultadoPapeis, Abs(vsValorOrig - vpValorCalc), 2, blachei, ValorTitulo, intIndex)
            
            If blachei = False Then
                Return False
            End If
        End If

        vsQtdeBase = ResultadoPapeis.Item(intIndex).Disponivel

        If ResultadoPapeis.Item(intIndex).TipoPrecoUnitario = "PU550" Then
            vsQtde = CType((vsValorOrig - vpValorCalc) / ResultadoPapeis. Item(intIndex).VL_Pu_550, Integer)
        Else
            vsQtde = CType((vsValororig - vpValorCalc) / ResultadoPapeis.Item(intIndex).VL_Pu_Mercado, Integer) 
        End If

        If vsQtde > vsQtdeBase Then
            vsQtde = vsQtdeBase
        End If
        
        If ResultadoPapeis.Item(intIndex).TipoPrecoUnitario = "PUS58" Then 
            vsValorLastro = vsQtde * ResultadoPapeis.Item(intIndex).VL_Pu_558
        Else
            vsValorLastro = vsQtde * ResultadoPapeis.Item(intIndex).VL_Pu_Mercado
        End If


        ResultadoPapeis.Item(intIndex).QTDE_Titulo = ResultadoPapeis.Item(intIndex).Disponivel 
        ResultadoPapeis.Item(intIndex).Disponivel = (ResultadoPapeis.Item(intIndex).Disponivel - vsQtde) 
        ResultadoPapeis.Item(intIndex).QTDE_Disponivel = ResultadoPapeis.Item(intIndex).Disponivel 
        ResultadoPapeis.Item(intIndex).QTDE_Utilizada = vsQtde
        If ResultadoPapeis.Item(intIndex).TipoPrecoUnitario = "PU550" Then
            ResultadoPapeis.Item(intIndex).Valor = Int(ResultadoPapeis.Item(intIndex).Disponivel * ResultadoPapeis.Item(intIndex).VI_Pu_550) 
        Else   
            ResultadoPapeis.Item(intIndex).Valor = Int(ResultadoPapeis.Item(intIndex).Disponivel * ResultadoPapeis.Item(intIndex).VI_Pu_Mercado) 
        End If
        ResultadoPapeis.Item(intIndex).TipoEstoque = vItemPapeis.ID_ESTOQUE
        
        PapeisUtilizados.Add(ResultadoPapeis.Item(intIndex))
    End While
    Return True
End Function

    Public Function CalularPUResgate (ByVal prDt_Inicio As Date,                   
                                        ByVal prDt_Fim As Date,
                                        ByVal prDT_Vencimento Papel As Date, 
                                        ByVal prCd_Indexador As String, 
                                        Byval prCd_Papel As Integer, 
                                        ByVal prPU550 As Double,
                                        Byval prVlTaxa As Double) As Double

        Dim Inicio As Date
        Dim SDtTermino As Date
        Dim intDifDias As Integer
        Dim sPrxVenctoPapel As Date
        Dim VenctoPapel As Date

        'Dim godatas As New GerenciadorDatas.Datas

        Dim vAcessoDados As New AcessoDados.Configuracao
        Dim dblVlPuida As Double
        Dim dblV1PuResg As Double
        Dim dblVlPu As Double
        Inicio = prDt_Inicio

        SDtTermino = prDt_Fim
        VenctoPapel =  prDT_Vencimento_Papel

        '-- Verifica a qtde de dias uteis entre a data de inicio e termino da operação
        '--------int DifDias = go_datas.difdu (CLng (SDt Inicio), CLng (SDtTermino)) 
        'TODO: Fazer comparação com o resultado do datas antigo int DifDias = Itau.MM.Framework.GerenciadorDatas.Datas.NumeroDiasUteis (Inicio.AddDays (1), SDtTermino)

        '-- cd indexador=PRE, dias uteis=1
        If (prcd_Indexador = "PRE") And intDifDias = 1 Then

        '-- A data de vencimento do papel é dia util
        '-- se não for procurar proximo dia util
        '---intDiaUtil – go_datas.IsUtilde (CLng (SDtVenctoPapel))

        If Not Itau.MM.Framework.GerenciadorDatas.Datas.DiaUtil(Vencto Papel.AddDays (1)) Then '--False = não é dia útil, True = é dia útil 
            SPrxVenctoPapel = Itau.MM.Framework.GerenciadorDatas.Datas.ObterProximoDiaUtil(Vencto Papel.AddDays (1)) 
        Else
            SPrxVenctoPapel = VenctoPapel

        End If

        '-- Dt Vencto da operação é igual a DT Vencto do papel 
        If SDtTermino = sPrxVenctoPapel Then
            Dim DadosPU550 As New Contrato.DadosPU550
            DadosPU550 = vAcessoDados.ConsultaPU5550(prcd_Papel)

            CalularPUResgate = 0

            '-- Aplicar a formula
            'dblVlPuida = Trunca (8, 1000 / (((CDbl(Troca (Txt_Rentabilidade.Text, ".", ",")) / 100 + 1)^ (Val (int DifDias) / 252)))) ,
            'dblV1PuResg = 1000
            dblVlPu = DadosPU550.VL_PU_RET
            dblvlPuida = Trunca(8, dblvipu / (((CDbl(Troca(prvlTaxa.ToString, ".", ",")) / 100 + 1) ^ (Val(intDifDias) / 252)))) 
            dblV1PuResg = dblVlPu
            CalularPUResgate = dblV1Puida


        Else
            dblVlPu = prPU550
            dblV1Puida = Trunca(8, dblVlPu * (((CDbl(Troca(prVlTaxa.ToString, ".", ",")) / 100 + 1) (Val(intDifDias) / 252))))
            dblVlPuResg = db1V1Puida
            CalularPUResgate = dblV1Puida
        End If

    End Function


    Public Sub ConfirmaValorPUna_Vespera_Vencto_Papel(ByVal prvi Pu Volta As Double, Byval prvi Pu 550 As Double, ByVal prvl_taxa As Double, ByVal pri 
        Dim rtVlPu As Double 
        Dim rtVlPuida As Double

        'verificando a conta reversa...baseado no PUIda calcular o PUResg...
        rtVlPu = prvi_Pu_550
        rtVlPuida = Trunca(8, rtvipu (((CDbl (prV1_taxa)/100+ 1)^(Val(prNU_prazo_du) / 252))))
        If rtVlPuida < prvl_Pu_Volta Then
            rtVlPu = rtVlPu + 0.00000001
            rtVlPuida = Trunca(8, rtVlPu* (((CDbl(prV1_taxa) / 100 + 1) ^ (Val(prNU_prazo_du) / 252)))) 
            If rtVlPuida = prVi_Pu_Volta Then
                rtVl_Pu_550 = rtVlPu
                rtVl_Pu_Volta = prVl_Pu_Volta
            Else
                rtVl_Pu_550 = prvl_Pu_550 
                rtVl_Pu_Volta = prV1_Pu_Volta 
            End If

        Else If rtVlPuida > prVl_Pu_Volta Then
            rtVlPu = rtvipu - 0.00000001
            rtVlPuida = Trunca (8, rtVlPu* (((CDbl(prV1_taxa) / 100 + 1)^ (Val (prNU_prazo_du) / 252)))) 
            If rtVlPuida = prVi_Pu_Volta Then
                rtVl_Pu_550 = rtVlPu
                rtVl_Pu_Volta = prV1_Pu_Volta
            Else
                rtVl_Pu_550 = prV1_Pu_550
                rtVl_Pu_Volta = prV1_Pu_Volta 
            End If
        Else
            rtVl_Pu_550 = prv1_Pu_550
            rtVl_Pu_Volta = prV1_Pu_Volta
        End If

    End Function


    Public Function Trunca (ByVal Num Dec As Integer, ByVal Numero As Double) As Double

        'Trunca = Int (Format((Numero 10 Num Dec))) / (10 Num Dec)
        Dim truncal As Double
        Dim trunca2 As Double
        Dim trunca3 As Double
        Dim trunca4 As Double

        'truncal Int (System.Convert.ToDecimal (Numero 10 Num_Dec))
        'trunca2 = (10Num Dec)
        'trunca3 = truncal / trunca2
        'Trunca trunca3

        truncal = Numero * 100 ^ Num_Dec
        trunca2 = Fix(truncal)
        trunca3 = trunca2 / 10 ^ Num_Dec
        trunca4 = Fix(trunca3)
        Trunca = trunca4/10 ^ Num_Dec
    End Function
        


    Public Function Troca (ByVal X As String, ByVal Charl As String, ByVal Char2 As String) As String 
        'Troca caracter "Charl" pelo caracter "Char2" na String "x"
        'Verifica se a string nao tem "Charl"

        Dim pos As Integer

        pos = InStr(1, X, Charl)

        If pos = 0 Then
            Troca = X
        Else
            While CBool(InStr(1, X, Charl))
                pos = InStr(1, X, Charl)
                X = Mid$(X, 1, pos- 1) + Char2 + Mid$(X, pos + 1, Len(X) = pos)
            End While
            Troca = X
        End If
    End Function



    Public Sub Grava_Complemento_Cotacao (ByVal vpDataMesa As Date, ByVal vpCdCotacao As String) 
        Dim strl As String = ""
        Dim vPar As String =
        Dim iContaErro As Integer = 0
        Dim bTentarNovamente As Boolean = False










































