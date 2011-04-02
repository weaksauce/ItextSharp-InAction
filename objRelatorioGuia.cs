using System;
using System.Drawing;
using System.IO;
using ObjetosDeNegocio;
using System.Web.SessionState;
using DesWebServer;
using System.Data;
using System.Collections;

namespace ObjetosDeNegocio
{
    public class GerarGuiaLote : TipoProcessamento
    {
        public ArrayList pc;
        public FormImpressaoGuiaLote formImpressaoGuiaLote;
        public ConvenioBanco cb;
        public C1.C1PrintDocument.C1PrintDocument repGuia;
        public string pathAplicacao;
        public Usuario usuario;
        DataTable dt;

        public GerarGuiaLote() { }

        public override string tipoProcessamento { get { return ""; } }
        public override int nivel { get { return 1; } }
        public override string periodicidade { get { return ""; } }

        public override bool CalcularTodos(G6Session g6Session)
        {
            GerarGuias(g6Session);
            return true;
        }

        private Image imgArquivoBrasao = null;
        private Image imgLogo = null;
        private Image imgArquivoLogotipoCorreio = null;

        public void GerarGuias(G6Session g6Session)
        {
            int qtdeGuias = pc.Count;
            int qtdeGuiasPorArquivo = Parametro.GetParametroInteiroPorNome("QtdeGuiasInserirArquivoPDF");
            if (qtdeGuiasPorArquivo <= 0)
                qtdeGuiasPorArquivo = 40;

            string filenamecompleto;

            Image img= null;
            try
            {
                img = Image.FromFile(Parametro.GetParametroPorNome("ArquivoBrasao"));
                //imgArquivoBrasao = Image.FromFile(Parametro.GetParametroPorNome("ArquivoBrasao"));
                imgArquivoBrasao = (Image)img.Clone();
                img.Dispose();
            }
            catch (Exception ex)
            {
                Utils.RegisterLogEvento("Problema ao carregar o arquivo de brasão no RelatórioGuia. Erro: " + ex.Message, "GuiaRecolhimento", Utils.EventType.Error, ex.StackTrace, null);
                //throw;
            }

            string pathLogoBanco = Parametro.GetParametroPorNome("ArquivoLogotipoBanco");
            if (!String.IsNullOrEmpty(pathLogoBanco))
            {
                img = Image.FromFile(pathLogoBanco);//"c:\\inetpub\\wwwroot\\deswebserver\\images\\logo_caixa.gif");
                imgLogo = (Image)img.Clone();
                img.Dispose();
            }

            string pathLogoCorreio = Parametro.GetParametroPorNome("ArquivoLogotipoCorreio");
            if (!String.IsNullOrEmpty(pathLogoCorreio))
            {
                try
                {
                    img = Image.FromFile(pathLogoCorreio);
                    imgArquivoLogotipoCorreio = (Image)img.Clone();
                    img.Dispose();
                }
                catch (Exception) { }
            }

            dt = new DataTable();
            dt.Columns.Add("Intervalo das Guias");
            dt.Columns.Add("Nome do Arquivo");

            int posOrig = 0;
            int pos = 1;

            while (posOrig < pc.Count)
            {
                GuiaRecolhimento[] guias;
                if (pc.Count > qtdeGuiasPorArquivo)
                {
                    if ((pc.Count - posOrig) >= qtdeGuiasPorArquivo)
                        guias = new GuiaRecolhimento[qtdeGuiasPorArquivo];
                    else
                        guias = new GuiaRecolhimento[(pc.Count - posOrig)];
                }
                else
                    guias = new GuiaRecolhimento[pc.Count];
                int posDest = 0;
                while (posDest < qtdeGuiasPorArquivo && posOrig < pc.Count)
                {
                    guias[posDest] = (GuiaRecolhimento)pc[posOrig];
                    posDest++;
                    posOrig++;
                }
                string sdt = DateTime.Now.ToString("yyyyMMddHHmmss");
                repGuia.DocumentName = "guias-" + guias[0].numGuia.ToString() + "-" + sdt + ".pdf";
                RelatorioGuia rel = new RelatorioGuia(ref repGuia, guias, cb);
                rel.imgArquivoBrasao = imgArquivoBrasao;
                rel.imgLogo = imgLogo;
                rel.imgArquivoLogotipoCorreio = imgArquivoLogotipoCorreio;

                rel.NomeUsuarioCorrente = usuario.nome;
                rel.PathAplicacao = pathAplicacao;
                rel.FrenteVerso = true;
                repGuia.StartDoc();
                InicioContadorProgresso(pc.Count);

                for (int i = 0; i < guias.Length; i++)
                {
                    rel.guia = guias[i];
                    if (rel.guia != null && rel.guia.valorGuia >= 0.01m)
                    {
                        rel.guia.Retrieve();
                        rel.DesenharCabecalhoPrefeitura();
                        rel.DesenharTitulo("DAM - Documento de Arrecadação Municipal");
                        rel.DesenharCabecalhoBanco();
                        rel.DesenharDadosGuia();
                        rel.DesenharDadosContribuinte();
                        rel.DesenharLancamentos(); // o desenho do boleto esta dentro da funcao de DesenharLancamentos
                        rel.NovaPagina();
                        //altera o status da guia para impressa
                        if (rel.guia.status == "G")
                            rel.guia.status = "N";
                        if (rel.guia.status != "E")
                            rel.guia.numEmissao++;

                        rel.guia.Save(g6Session != null ? g6Session.persistentTransaction : null);
                        //RF015 - Registro de LOG Geração PDF Impressão Guia (em lote)   
                        EventoSeguranca.CriarEventoSeguranca(TipoEventoSeguranca.GetTipoEventoSegurancaById("GeracaoGuia"),
                            "Geração de PDF para impressão da Guia (em lote)", "", rel.guia.oid, 0, 0, g6Session);

                        //if (i < (guias.Length - 2))
                        //{
                        IncrementarContadorProgresso(pos);
                        pos++;
                        //}
                        //Session["GuiaAtual"] = (i+1).ToString();
                    }
                }

                repGuia.EndDoc();

                if (!Directory.Exists(pathAplicacao + "\\" + Parametro.GetParametroPorNome("NomeDiretorioGeraGuias")))
                    Directory.CreateDirectory(pathAplicacao + "\\" + Parametro.GetParametroPorNome("NomeDiretorioGeraGuias"));

                filenamecompleto = pathAplicacao + "\\" + Parametro.GetParametroPorNome("NomeDiretorioGeraGuias") + "\\" + repGuia.DocumentName;
                repGuia.ExportToPDF(filenamecompleto, true);

                string intervalo = "Entre " + guias[0].numGuia.ToString() + " até " + guias[guias.Length - 1].numGuia.ToString();
                dt.Rows.Add(new string[] { intervalo, repGuia.DocumentName });
            }

            if (imgArquivoBrasao != null)
                imgArquivoBrasao.Dispose();
            if (imgLogo != null)
                imgLogo.Dispose();
            if (imgArquivoLogotipoCorreio != null)
                imgArquivoLogotipoCorreio.Dispose();
        }

        public DataTable GetGridGuias()
        {
            return dt;
        }

    }

	/// <summary>
	/// Classe responsável pela criação do boleto bancario para impressao
	/// </summary>
	public class RelatorioGuia : G6RelatorioManual
	{
        public Image imgArquivoBrasao = null;
        public Image imgLogo = null;
        public Image imgArquivoLogotipoCorreio = null;

		bool frenteVerso;
		public bool FrenteVerso
		{
			set
			{
				frenteVerso = value;
			}
		}

		string nomeUsuario="Sistema";
		public string NomeUsuarioCorrente
		{
			set 
			{
				nomeUsuario = value;
			}
		}
		

		public GuiaRecolhimento guia;

		float alturaBoleto = 16.78f;

		//float alturaLancamentosComBoleto = 8.8f;
		int numLinhasLancamentoComBoleto = 22;
		int numCaracteresHistoricoLancamento = 71;
		int numMaxCaracteresHistoricoLancamentoComBoleto;
		
		float cabLancHist=2.5f;
		float cabLancVenc=14.3f;
		float cabLancValor=17f;

        EstruturaBoletoBancario ebb = null;
		ConvenioBanco cb;
		GuiaRecolhimento[] guiasRecolhimento;

		public RelatorioGuia(ref C1.C1PrintDocument.C1PrintDocument _rep, ref GuiaRecolhimento _guia, ConvenioBanco _cb) : base(ref _rep)
		{			
			rep = _rep;
			guia =_guia;
			guia.Retrieve();
			cb = _cb;
			cb.Retrieve();

            ebb = EstruturaBoletoBancario.GetEstruturaBoleto(cb.banco.codBanco);

			rep.C1DPageSettings = "color:False;landscape:False;margins:15,15,15,15;papersize:827,1169,QQA0AA==";
			rep.ColumnSpacingStr = "1.26999998092651cm";
			rep.ColumnSpacingUnit.DefaultType = true;
			rep.ColumnSpacingUnit.UnitValue = "1.26999998092651cm";
			rep.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Cm;
			rep.DefaultUnitOfFrames = C1.C1PrintDocument.UnitTypeEnum.Cm;
			rep.DefaultUnitOfFramesSynchronized = false;
		}

		public RelatorioGuia(ref C1.C1PrintDocument.C1PrintDocument _rep, GuiaRecolhimento[] _guiasRecolhimento, ConvenioBanco _cb) : base(ref _rep)
		{			
			rep = _rep;
			guiasRecolhimento = _guiasRecolhimento;

			cb = _cb;
			cb.Retrieve();

            ebb = EstruturaBoletoBancario.GetEstruturaBoleto(cb.banco.codBanco);

			rep.C1DPageSettings = "color:False;landscape:False;margins:15,15,15,15;papersize:827,1169,QQA0AA==";
			rep.ColumnSpacingStr = "1.26999998092651cm";
			rep.ColumnSpacingUnit.DefaultType = true;
			rep.ColumnSpacingUnit.UnitValue = "1.26999998092651cm";
			rep.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Cm;
			rep.DefaultUnitOfFrames = C1.C1PrintDocument.UnitTypeEnum.Cm;
			rep.DefaultUnitOfFramesSynchronized = false;

		}

        public void GerarRelatorio(G6Session g6Session)
		{
            Image img = null;
            if (imgArquivoBrasao == null)
            {
                try
                {
                    img = Image.FromFile(Parametro.GetParametroPorNome("ArquivoBrasao"));
                    //imgArquivoBrasao = Image.FromFile(Parametro.GetParametroPorNome("ArquivoBrasao"));
                    imgArquivoBrasao = (Image)img.Clone();
                    img.Dispose();
                }
                catch (Exception ex)
                {
                    Utils.RegisterLogEvento("Problema ao carregar o arquivo de brasão no RelatórioGuia. Erro: " + ex.Message, "GuiaRecolhimento", Utils.EventType.Error, ex.StackTrace, null);
                    //throw;
                }
            }

            if (imgLogo == null)
            {
                string pathLogoBanco = Parametro.GetParametroPorNome("ArquivoLogotipoBanco");
                if (!String.IsNullOrEmpty(pathLogoBanco))
                {
                    img = Image.FromFile(pathLogoBanco);//"c:\\inetpub\\wwwroot\\deswebserver\\images\\logo_caixa.gif");
                    imgLogo = (Image)img.Clone();
                    img.Dispose();
                }
            }

            if(imgArquivoLogotipoCorreio == null)
            {
                string pathLogoCorreio = Parametro.GetParametroPorNome("ArquivoLogotipoCorreio");
                if (!String.IsNullOrEmpty(pathLogoCorreio))
                {
                    try
                    {
                        img = Image.FromFile(pathLogoCorreio);
                        imgArquivoLogotipoCorreio = (Image)img.Clone();
                        img.Dispose();
                    }
                    catch (Exception) { }
                }
            }

			DesenharCabecalhoPrefeitura();
			DesenharTitulo("DAM - Documento de Arrecadação Municipal");  
			DesenharCabecalhoBanco();
			DesenharDadosGuia();
			DesenharDadosContribuinte();
			DesenharLancamentos(); // o desenho do boleto esta dentro da funcao de DesenharLancamentos
			//altera o status da guia para impressa
			//guia.status = "N";
			//guia.numEmissao++;
            if (guia.status == "G")
                guia.status = "N";
            if (guia.status != "E")
                guia.numEmissao++;
			guia.Save();

            if (imgArquivoBrasao != null)
                imgArquivoBrasao.Dispose();
            if (imgLogo != null)
                imgLogo.Dispose();
            if (imgArquivoLogotipoCorreio != null)
                imgArquivoLogotipoCorreio.Dispose();

            //RF015 - Registro de LOG Geração/Impressão Guia    
            EventoSeguranca.CriarEventoSeguranca(TipoEventoSeguranca.GetTipoEventoSegurancaById("GeracaoGuia"),
                "Geração de PDF para impressão da Guia", "", guia.oid, 0, 0, g6Session);
		}

		public int GerarRelatorioVariasGuias(G6Session g6Session)
		{
			for(int i=0; i < guiasRecolhimento.Length; i++)
			{
				guia = guiasRecolhimento[i];
				if(guia != null && guia.valorGuia >= 0.01m)
				{
					guia.Retrieve();
					DesenharCabecalhoPrefeitura();
					DesenharTitulo("DAM - Documento de Arrecadação Municipal");  
					DesenharCabecalhoBanco();
					DesenharDadosGuia();
					DesenharDadosContribuinte();
					DesenharLancamentos(); // o desenho do boleto esta dentro da funcao de DesenharLancamentos
					rep.NewPage();
                    //altera o status da guia para impressa
                    if (guia.status == "G")
                        guia.status = "N";
                    if (guia.status != "E")
                        guia.numEmissao++;
					guia.Save();

                    //RF015 - Registro de LOG Geração/Impressão Guia    
                    EventoSeguranca.CriarEventoSeguranca(TipoEventoSeguranca.GetTipoEventoSegurancaById("GeracaoGuia"),
                        "Geração da Guia", "", guia.oid, 0, 0, g6Session);

                    if (i < (guiasRecolhimento.Length - 2))
                        return (i + 1);
						//Session["GuiaAtual"] = (i+1).ToString();
				}
			}
            return 0;
		}

		/// <summary>
		/// Desenha o retangulo do cabeçalho do boleto onde contem informacoes da PM
		/// </summary>
		public void DesenharCabecalhoPrefeitura()
		{
			//DesenharRetagulo(pontoXInicial, pontoYInicial, pontoXFinal,2.5f); 
			float posxx = pontoXInicial;
			float posyy = pontoYInicial+0.1f;
			string nomePM = Parametro.GetParametroPorNome("NomePrefeitura");
			string secretariaPM = Parametro.GetParametroPorNome("NomeSecretaria");
			string nomeCoordenadoria = Parametro.GetParametroPorNome("NomeCoordenadoria");
			string enderecoCompletoPM = Parametro.GetParametroPorNome("EnderecoCompleto");
			string fone = Parametro.GetParametroPorNome("Fone");
			if(!fone.Trim().Equals(""))
				enderecoCompletoPM += " Tel.: "+fone;

			float posesqtexto = posicaoEsquerda;

			//C:\\Inetpub\\wwwroot\\DesWebServer\\images
            //Image imgArquivoBrasao = null;
            //try
            //{
            //    imgArquivoBrasao = Image.FromFile(Parametro.GetParametroPorNome("ArquivoBrasao"));
            //}
            //catch(Exception ex)
            //{
            //    Utils.RegisterLogEvento("Problema ao carregar o arquivo de brasão no RelatórioGuia. Erro: " + ex.Message, "GuiaRecolhimento", Utils.EventType.Error, ex.StackTrace, null);
            //    //throw;
            //}
			C1.C1PrintDocument.ImageAlignDef ia = new C1.C1PrintDocument.ImageAlignDef();
			//20mm
            if (imgArquivoBrasao != null)
            {
                if (imgArquivoBrasao.Width < (2 * imgArquivoBrasao.Height))
                {
                    rep.RenderDirectImage(RetornaStringPosicao(posicaoEsquerda + 0.1f), RetornaStringPosicao(posyy + 0.05f), imgArquivoBrasao, 1.5f, 1.5f, ia);
                    posesqtexto = posesqtexto + 2.1f;
                }
                else
                {
                    rep.RenderDirectImage(RetornaStringPosicao(posicaoEsquerda + 0.1f), RetornaStringPosicao(posyy + 0.05f), imgArquivoBrasao, 4f, 1.5f, ia);
                    posesqtexto = posesqtexto + 4.5f;
                }
                //imgArquivoBrasao.Dispose();
            }

			Font f = new Font("Tahoma", 12, FontStyle.Bold);
			float w = GetComprimentoStringCentrimentros(f, nomePM);
			lastLinha += 0.07f;
			rep.RenderDirectText(RetornaStringPosicao(posesqtexto), RetornaStringPosicao(posyy),
				nomePM, RetornaStringPosicao(w), f, Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);

			lastLinha += GetAlturaStringCentrimentros(f, nomePM);
			posyy += + GetAlturaStringCentrimentros(f, nomePM);

			f = new Font("Tahoma", 10);
			w = GetComprimentoStringCentrimentros(f, secretariaPM);
			rep.RenderDirectText(posesqtexto, posyy, secretariaPM, w, f, Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);
			
			lastLinha += GetAlturaStringCentrimentros(f, secretariaPM);
			posyy += GetAlturaStringCentrimentros(f, secretariaPM);

			f = new Font("Tahoma", 9);
			w = GetComprimentoStringCentrimentros(f, nomeCoordenadoria);
			rep.RenderDirectText(posesqtexto, posyy, nomeCoordenadoria, w, f, Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);
			
			lastLinha += GetAlturaStringCentrimentros(f, nomeCoordenadoria);
			posyy += GetAlturaStringCentrimentros(f, nomeCoordenadoria);

			
			f = new Font("Tahoma", 8);
			w = GetComprimentoStringCentrimentros(f,enderecoCompletoPM); 
			rep.RenderDirectText(posesqtexto, posyy, enderecoCompletoPM, w, f, Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);

            string da = "RECIBO DO SACADO";
            f = new Font("Tahoma", 9);
            w = GetComprimentoStringCentrimentros(f, da);
            rep.RenderDirectText(RetornaStringPosicao(pontoXFinal - w), RetornaStringPosicao(posyy), da,
                GetComprimentoStringCentrimentros(f, da), f, Color.Gray, AlinharDireita());

			lastLinha = 2.5f; //altrua do retangulo
		}

		/// <summary>
		/// Desenha o retangulo do titulo do boleto
		/// </summary>
		/// <param name="titulo">Texto que sera apresentado como titulo</param>
		public void DesenharTitulo(string titulo)
		{
			lastLinha += 0.1f;
			rep.RenderDirectRectangle(RetornaStringPosicao(pontoXInicial), lastLinha,
				RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(lastLinha+0.5f),
				Color.Gray, 0.1f, Color.Gray);
			Font f = new Font("Arial", 12, FontStyle.Bold);
			float w = GetComprimentoStringCentrimentros(f, titulo);
			rep.RenderDirectText(RetornaStringPosicao(pontoXInicial+4.5f), lastLinha+0.02f, 
				titulo,w, f, Color.White, C1.C1PrintDocument.AlignHorzEnum.Center); 			
			lastLinha += 0.5f;
		}

		/// <summary>
		/// Cabecalho do banco
		/// </summary>
		public void DesenharCabecalhoBanco()
		{
			lastLinha += 0.1f;
			float altura = lastLinha+0.8f;
			DesenharRetagulo(pontoXInicial, lastLinha, pontoXFinal, altura);

			//linha logo banco
			DesenharLinhaVertical(4, lastLinha, altura);

            if (cb.isProxy)
                cb.Retrieve();
            string banco = cb.banco.nome.Trim();
            string texto = banco.PadLeft(banco.Length < 14 ? banco.Length / 2 + 7 : 0, ' ');

            string pathLogoBanco = Parametro.GetParametroPorNome("ArquivoLogotipoBanco");
            if (pathLogoBanco != null && pathLogoBanco != "")
            {
                try
                {
                    //Image imgLogo = Image.FromFile(pathLogoBanco);//"c:\\inetpub\\wwwroot\\deswebserver\\images\\logo_caixa.gif");
                    C1.C1PrintDocument.ImageAlignDef align = new C1.C1PrintDocument.ImageAlignDef();
                    align.StretchHorz = true; align.StretchVert = true;
                    align.AlignHorz = C1.C1PrintDocument.ImageAlignHorzEnum.Center; align.AlignVert = C1.C1PrintDocument.ImageAlignVertEnum.Center;
                    rep.RenderDirectImage(posicaoEsquerda, RetornaStringPosicao(lastLinha + 0.05f), imgLogo, 3.5f, 0.6f, align);
                    //imgLogo.Dispose();
                }
                catch (Exception ex)
                {
                    Utils.RegisterLogEvento("Problema ao carregar o arquivo de Logotipo do Banco no RelatórioGuia. Erro: " + ex.Message, "GuiaRecolhimento", Utils.EventType.Error, ex.StackTrace, null);

                    Font f = new Font("Arial", 12, FontStyle.Bold);
                    float w = GetComprimentoStringCentrimentros(f, texto);
                    rep.RenderDirectText(posicaoEsquerda, RetornaStringPosicao(lastLinha + 0.2f), texto, w, f,
                        Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);
                }
            }
            else
            {
                Font f = new Font("Arial", 12, FontStyle.Bold);
			    float w = GetComprimentoStringCentrimentros(f, texto);
			    rep.RenderDirectText(posicaoEsquerda, RetornaStringPosicao(lastLinha+0.2f), texto, w, f,
                    Color.Black, C1.C1PrintDocument.AlignHorzEnum.Left);
            }
			//linha banco
			DesenharLinhaVertical(7, lastLinha, altura); 
			
			string codBanco = cb.banco.codBanco+"-"+cb.banco.digito;

			DesenharDescricaoCampo(4, lastLinha, "Banco", codBanco);

			string agenciacodigocedente = ebb.CalculoAgenciaCedente(cb.codAgencia, cb.digitoAgencia, cb.numConvenioLider, cb.numConta, cb.digitoConta);
			//agenciacodigocedente = cb.codAgencia.Trim()+"-"+cb.digitoAgencia.Trim()+"/"+cb.numConta.Trim()+"-"+cb.digitoConta.Trim();
			DesenharDescricaoCampo(7, lastLinha, "Agência Código Cedente", agenciacodigocedente);

            string nossonumero;
            nossonumero = ebb.CalculoNossoNumeroFormatado(guia.numGuia.ToString() + (guia.numParcela.ToString().PadLeft(2, '0')), cb);//.numConvenioLider, cb.codCarteira);

            DesenharLinhaVertical(11, lastLinha, altura);
            DesenharDescricaoCampo(11, lastLinha, "Nosso Número", nossonumero);

            DesenharLinhaVertical(15, lastLinha, altura);
            DesenharDescricaoCampo(15, lastLinha, "Vencimento", guia.dataVencimento.ToString("dd/MM/yyyy"));
            lastLinha = altura;
            
		}

		/// <summary>
		/// Retangulo com informacoes da guia de recolhimento
		/// </summary>
		public void DesenharDadosGuia()
		{
			DesenharRetagulo(pontoXInicial, lastLinha, pontoXFinal, lastLinha+alturaLinha);
			//linua numguia
			float yfim = lastLinha+alturaLinha;
			DesenharLinhaVertical(4, lastLinha, yfim);

			rep.RenderDirectRectangle(RetornaStringPosicao(posicaoEsquerda+0.02f), RetornaStringPosicao(lastLinha+0.07f), RetornaStringPosicao(3.9f), RetornaStringPosicao(yfim-0.1f), Color.Transparent, 0.1f, Color.LightGray);
			DesenharDescricaoCampo(posicaoEsquerda, lastLinha, "Nº Guia", guia.numGuia.ToString());
			
			//linha numparcela
			DesenharLinhaVertical(5.2f, lastLinha, yfim);
			rep.RenderDirectRectangle(RetornaStringPosicao(4.1f), RetornaStringPosicao(lastLinha+0.07f), RetornaStringPosicao(5.1f), RetornaStringPosicao(yfim-0.1f), Color.Transparent, 0.1f, Color.LightGray);
			DesenharDescricaoCampo(4, lastLinha, "Parcela", guia.numParcela.ToString("00"));
			
			//linha data emissao
			DesenharLinhaVertical(8, lastLinha, yfim);
			DesenharDescricaoCampo(5.2f, lastLinha, "Data de Emissão", DateTime.Today.ToString("dd/MM/yyyy"));
			
			//linha num emissao
			DesenharLinhaVertical(11, lastLinha, yfim);
			DesenharDescricaoCampo(8, lastLinha, "Nº Emissão", guia.numEmissao.ToString("00"));

			//operador
			DesenharDescricaoCampo(11, lastLinha, "Operador", nomeUsuario);
			lastLinha += alturaLinha;
		}

		/// <summary>
		/// Retangulo que apresenta os dados do contribuinte
		/// </summary>
		public void DesenharDadosContribuinte()
		{
			guia.contribuinte.Retrieve();
			lastLinha += 0.1f;
			float yfim = lastLinha+alturaLinha;
			//primeira Linha
			DesenharRetagulo(pontoXInicial, lastLinha, pontoXFinal, yfim );
			//DesenharLinhaVertical((pontoXFinal-pontoXInicial)/2, lastLinha, yfim);
			DesenharDescricaoCampo(posicaoEsquerda, lastLinha, "Razão Social", guia.contribuinte.nome);
			//DesenharDescricaoCampo((pontoXFinal-pontoXInicial)/2, lastLinha, "Nome Fantasia", guia.contribuinte.nomeFantasia);
			
			lastLinha = yfim;
			yfim = lastLinha + alturaLinha;
			//segunda linha
			DesenharRetagulo(pontoXInicial, lastLinha, pontoXFinal, yfim);
			//linha cnpj
			DesenharLinhaVertical(4, lastLinha, yfim);
			lastLinha -= 0.05f; 
			rep.RenderDirectRectangle(RetornaStringPosicao(posicaoEsquerda+0.02f), RetornaStringPosicao(lastLinha+0.07f), RetornaStringPosicao(3.9f), RetornaStringPosicao(yfim-0.05f), Color.Transparent, 0.1f, Color.LightGray);
			DesenharDescricaoCampo(posicaoEsquerda, lastLinha, "Cadastro Mobiliário", guia.contribuinte.inscricao.Trim());
			DesenharDescricaoCampo(4, lastLinha, "CNPJ/CPF", guia.contribuinte.numDocReceita);
			//linha fone
			DesenharLinhaVertical(8, lastLinha, yfim);
			if(guia.contribuinte.telefone.Count > 0)
				DesenharDescricaoCampo(8, lastLinha, "Fone", guia.contribuinte.telefone[0].ToString());
			else
				DesenharDescricaoCampo(8, lastLinha, "Fone", "");
			//linha email
			DesenharLinhaVertical(12, lastLinha, yfim);
			DesenharDescricaoCampo(12, lastLinha, "E-Mail", guia.contribuinte.email);
			lastLinha = yfim;
		}

		/// <summary>
		/// Apresenta o cabelho dos lancamentos
		/// </summary>
		private void DesenharCabecalhoLancamento()
		{
			numMaxCaracteresHistoricoLancamentoComBoleto = numLinhasLancamentoComBoleto * numCaracteresHistoricoLancamento;
			lastLinha += 0.1f;
			rep.RenderDirectRectangle(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(lastLinha),
				RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(lastLinha+alturaLinha),
				Color.Black, 0.1f, Color.Gray);
			
			Font f = new Font("Arial", 9, FontStyle.Bold);
			float w=GetComprimentoStringCentrimentros(f, "Data Lanc.");
			//linha historico
			DesenharLinhaVertical(cabLancHist, lastLinha, lastLinha+alturaLinha);
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda-0.2f), RetornaStringPosicao(lastLinha+0.2f), "Data Lanc.",
				RetornaStringPosicao(w), f, Color.Black, AlinharCentro());
			//linha data venc
			w = GetComprimentoStringCentrimentros(f, "Histórico");
			DesenharLinhaVertical(cabLancVenc, lastLinha, lastLinha+alturaLinha);
			rep.RenderDirectText(RetornaStringPosicao(7f), RetornaStringPosicao(lastLinha+0.2f), "Histórico",
				RetornaStringPosicao(w), f, Color.Black, AlinharCentro());

			//linha valor
			w = GetComprimentoStringCentrimentros(f, "Data Venc.");
			DesenharLinhaVertical(cabLancValor, lastLinha, lastLinha+alturaLinha);
			rep.RenderDirectText(RetornaStringPosicao(14.6f), RetornaStringPosicao(lastLinha+0.2f), "Data Venc.",
				RetornaStringPosicao(w), f, Color.Black, AlinharCentro());

			w = GetComprimentoStringCentrimentros(f, "Valor");
			rep.RenderDirectText(RetornaStringPosicao(17.8f), RetornaStringPosicao(lastLinha+0.2f), "Valor",
				RetornaStringPosicao(w), f, Color.Black, AlinharCentro());
			lastLinha += alturaLinha;
		}

		/// <summary>
		/// Gera os lancamentos da guia de recolhimento
		/// o boleto bancario será apresentado somente na ultima folha.
		/// se a opcao de frenteVerso estiver habilitada entao o sistema irá imprimir o verso do boleto.
		/// </summary>
		public void DesenharLancamentos()
		{
			DesenharCabecalhoLancamento();
			float linhaCabecalho = lastLinha;
			lastLinha += 0.1f;

			//	int numLinhas = Queries.GetNumLinhasHistoricoLancamento(numCaracteresHistoricoLancamento, guia.oid);
			Font f;

			string hist="";
			int lin=0;
			float w;
			f = new Font("Arial", 8);
			bool bSegVia = true;
            bool bLancMaisDeUmaFolha = false;
			foreach(Lancamento lanc in guia.lancamentos)
			{
				if(guia.numEmissao > 1 && bSegVia)
				{
					f = new Font("Arial", 100);
					string segvia = "2ª VIA";
					rep.RenderDirectText(RetornaStringPosicao(3), RetornaStringPosicao(12),segvia,
						GetComprimentoStringCentrimentros(f, segvia), f, Color.LightGray, AlinharEsquerda());
					f = new Font("Arial", 8);
					bSegVia = false;
				}

				lanc.Retrieve();
				hist = lanc.tipoLancamento.descricao+"/"+lanc.historico;
				w = GetComprimentoStringCentrimentros(f, hist);
				lin = (int)Math.Ceiling((w/11));
				if(lin <= 0)
					lin = 1;

                System.Text.RegularExpressions.Regex x = new System.Text.RegularExpressions.Regex("\n");
                lin += x.Matches(hist).Count;

				//data emissao
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+0.1f), RetornaStringPosicao(lastLinha),
					lanc.dataLancamento.ToString("dd/MM/yyyy"), GetComprimentoStringCentrimentros(f, lanc.dataLancamento.ToString("dd/MM/yyyy")),
					f, Color.Black, AlinharEsquerda());
				//historico
				rep.RenderDirectText(RetornaStringPosicao(2.55f), RetornaStringPosicao(lastLinha), 
					hist, RetornaStringPosicao(11.45f), f, Color.Black, AlinharEsquerda());
				//data vencimento
				rep.RenderDirectText(RetornaStringPosicao(15f), RetornaStringPosicao(lastLinha),
					lanc.dataVencimento.ToString("dd/MM/yyyy"), GetComprimentoStringCentrimentros(f, lanc.dataVencimento.ToString("dd/MM/yyyy")),
					f, Color.Black, AlinharEsquerda());
				//valor
				String vlr="";
				if(lanc.valor > 1000)
						vlr = lanc.valor.ToString("0,000.00");
				else
					vlr = lanc.valor.ToString("0.00");
				vlr = vlr.PadLeft(10, ' ');

					w = GetComprimentoStringCentrimentros(f, vlr);
					rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f), RetornaStringPosicao(lastLinha),
						vlr, GetComprimentoStringCentrimentros(f, vlr),
						f, Color.Black, AlinharDireita());
				//}
				//else
				//{
				//	w = GetComprimentoStringCentrimentros(f, lanc.valor.ToString("0.00"));
				//	rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f), RetornaStringPosicao(lastLinha),
				//		lanc.valor.ToString("0.00"), GetComprimentoStringCentrimentros(f, lanc.valor.ToString("0.00")),
				//		f, Color.Black, AlinharDireita());
				//}

				lastLinha += (0.4f * lin); 
				if(lastLinha > 26)
				{
					//laterais
					DesenharLinhaVertical(pontoXInicial, linhaCabecalho, lastLinha);
					DesenharLinhaVertical(pontoXFinal, linhaCabecalho, lastLinha);
					//historico
					DesenharLinhaVertical(cabLancHist, linhaCabecalho, lastLinha);
					//vencimento
					DesenharLinhaVertical(cabLancVenc, linhaCabecalho, lastLinha);
					DesenharLinhaVertical(cabLancValor, linhaCabecalho, lastLinha);
					//abaixo
					rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(lastLinha),
						RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(lastLinha));

					DesenharVerso();
					rep.NewPage();
                    bLancMaisDeUmaFolha = true;
					lastLinha = 0.5f;
					DesenharCabecalhoPrefeitura();
					DesenharTitulo("DAM - Documento de Arrecadação Municipal");  
					DesenharCabecalhoBanco();
					DesenharDadosGuia();
					DesenharDadosContribuinte();
					DesenharCabecalhoLancamento();
					bSegVia = true;
				}
			}//foreach

			if(lastLinha > alturaBoleto)
			{
				DesenharVerso();
				rep.NewPage();
                bLancMaisDeUmaFolha = true;
				lastLinha = 0.5f;
				DesenharCabecalhoPrefeitura();
				DesenharTitulo("DAM - Documento de Arrecadação Municipal");  
				DesenharCabecalhoBanco();
				DesenharDadosGuia();
				DesenharDadosContribuinte();
				DesenharCabecalhoLancamento();
				//laterais
				DesenharLinhaVertical(pontoXInicial, linhaCabecalho, alturaBoleto);
				DesenharLinhaVertical(pontoXFinal, linhaCabecalho, alturaBoleto);
				//historico
				DesenharLinhaVertical(cabLancHist, linhaCabecalho, alturaBoleto);
				//vencimento
				DesenharLinhaVertical(cabLancVenc, linhaCabecalho, alturaBoleto);
				DesenharLinhaVertical(cabLancValor, linhaCabecalho, alturaBoleto);
				//abaixo
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(alturaBoleto),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(alturaBoleto));
				if(guia.numEmissao > 1 && bSegVia)
				{
					f = new Font("Arial", 100);
					string segvia = "2ª VIA";
					rep.RenderDirectText(RetornaStringPosicao(3), RetornaStringPosicao(12),segvia,
						GetComprimentoStringCentrimentros(f, segvia), f, Color.LightGray, AlinharEsquerda());
					f = new Font("Arial", 8);
					bSegVia = false;
				}
                DesenharBoleto(bLancMaisDeUmaFolha);
			}
			else
			{
				//laterais
				DesenharLinhaVertical(pontoXInicial, linhaCabecalho, alturaBoleto);
				DesenharLinhaVertical(pontoXFinal, linhaCabecalho, alturaBoleto);
				//historico
				DesenharLinhaVertical(cabLancHist, linhaCabecalho, alturaBoleto);
				//vencimento
				DesenharLinhaVertical(cabLancVenc, linhaCabecalho, alturaBoleto);
				DesenharLinhaVertical(cabLancValor, linhaCabecalho, alturaBoleto);
				//abaixo
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(alturaBoleto),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(alturaBoleto));
				if(guia.numEmissao > 1 && bSegVia)
				{
					f = new Font("Arial", 100);
					string segvia = "2ª VIA";
					rep.RenderDirectText(RetornaStringPosicao(3), RetornaStringPosicao(12),segvia,
						GetComprimentoStringCentrimentros(f, segvia), f, Color.LightGray, AlinharEsquerda());
					f = new Font("Arial", 8);
					bSegVia = false;
				}
                DesenharBoleto(bLancMaisDeUmaFolha);
			}
			DesenharVerso();

		}

        public void NovaPagina()
        {
            rep.NewPage();
        }

		/// <summary>
		/// Desenha o boleto bancario em si -- boleto do banco
		/// </summary>
		private void DesenharBoleto(bool bEscreverLancamentosEmAnexo)
		{
			float linhaDireita = 14;
			float linha= alturaBoleto;
			Font f = new Font("Arial", 9, FontStyle.Bold);
			string da;
			float w, h, altura;

            if(bEscreverLancamentosEmAnexo)
                DesenharDescricaoCampo(cabLancHist, linha - 0.7f, "", "Descrição dos lançamentos em anexo.");

			string obs = Parametro.GetParametroPorNome("ObservacaoLancamentoGuia");

			//retangulo obs
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, linha+alturaLinha);
			DesenharLinhaVertical(cabLancVenc, linha, linha+alturaLinha);
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Obs.", obs);
			da = "Total em R$";
			w = GetComprimentoStringCentrimentros(f, da);
			rep.RenderDirectText(RetornaStringPosicao(cabLancVenc+0.2f), RetornaStringPosicao(linha+0.2f), 
				da, RetornaStringPosicao(w), f, Color.Black, AlinharEsquerda());
			if(guia.valorGuia > 1000)
			{
				da = guia.valorGuia.ToString("0,000.00");
				w = GetComprimentoStringCentrimentros(f, da);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f), RetornaStringPosicao(linha+0.2f), 
					da, RetornaStringPosicao(w), f, Color.Black, AlinharDireita());			
			}
			else
			{
				da = guia.valorGuia.ToString("0.00");
				w = GetComprimentoStringCentrimentros(f, da);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f), RetornaStringPosicao(linha+0.2f), 
					da, RetornaStringPosicao(w), f, Color.Black, AlinharDireita());	
			}

			f = new Font("Tahoma", 9);
			linha += alturaLinha;

			da = "Autenticação Mecânica";
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda), RetornaStringPosicao(linha), da,
				GetComprimentoStringCentrimentros(f, da), f, Color.Gray , AlinharEsquerda());

			da = "RECIBO DO SACADO";
			w = GetComprimentoStringCentrimentros(f,da); 
			//rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w), RetornaStringPosicao(linha), da,
			//	GetComprimentoStringCentrimentros(f, da), f, Color.Gray , AlinharDireita());

			h = GetAlturaStringCentrimentros(f, da);
			linha += h;

			//desenho da linha de separacao
			da = "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"; 
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda), RetornaStringPosicao(linha), da,
				GetComprimentoStringCentrimentros(fontDescricao, da), fontDescricao, Color.Gray , AlinharEsquerda());
			h = GetAlturaStringCentrimentros(fontDescricao, da);
			linha += h;
			f = new Font("Tahoma", 10);
			da = "FICHA DE COMPENSAÇÃO";
			w = GetComprimentoStringCentrimentros(f,da); 
			rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w), RetornaStringPosicao(linha), da, RetornaStringPosicao(w), f,
				Color.Black, AlinharDireita());

			h = GetAlturaStringCentrimentros(f, da);
			linha += h; 
			altura= linha+1.1f;
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, altura);

			//linha logo banco
			DesenharLinhaVertical(4, linha, altura);

            string banco = cb.banco.nome.Trim();
            string texto = banco.PadLeft(banco.Length < 14 ? banco.Length / 2 + 7 : 0, ' ');

            string pathLogoBanco = Parametro.GetParametroPorNome("ArquivoLogotipoBanco");
            if (pathLogoBanco != null && pathLogoBanco != "")
            {
                try
                {
                    //Original
                    //Image imgLogo = Image.FromFile(pathLogoBanco);
                    //rep.RenderDirectImage(posicaoEsquerda, RetornaStringPosicao(linha + 0.1f), imgLogo);
                    //imgLogo.Dispose();

                    //Replicado do primeiro logo (DAS)
                    //Image imgLogo = Image.FromFile(pathLogoBanco);
                    C1.C1PrintDocument.ImageAlignDef align = new C1.C1PrintDocument.ImageAlignDef();
                    //align.StretchHorz = true; align.StretchVert = true;
                    align.AlignHorz = C1.C1PrintDocument.ImageAlignHorzEnum.Center; 
                    align.AlignVert = C1.C1PrintDocument.ImageAlignVertEnum.Center;
                    rep.RenderDirectImage(posicaoEsquerda, RetornaStringPosicao(linha + 0.2f), imgLogo, 3.5f, 0.6f, align);
                    //imgLogo.Dispose();

                }
                catch (Exception ex)
                {
                    Utils.RegisterLogEvento("Problema ao carregar o arquivo de Logotipo do Banco no RelatórioGuia. Erro: " + ex.Message, "GuiaRecolhimento", Utils.EventType.Error, ex.StackTrace, null);

                    f = new Font("Arial", 12, FontStyle.Bold);
                    w = GetComprimentoStringCentrimentros(f, texto);
                    rep.RenderDirectText(posicaoEsquerda, RetornaStringPosicao(linha + 0.3f), texto, w, f,
                        Color.Black, AlinharEsquerda());
                }
            }
            else
            {
                f = new Font("Arial", 12, FontStyle.Bold);
                w = GetComprimentoStringCentrimentros(f, texto);
                rep.RenderDirectText(posicaoEsquerda, RetornaStringPosicao(linha + 0.3f), texto, w, f,
                    Color.Black, AlinharEsquerda());
            }

            //linha banco
			DesenharLinhaVertical(7, linha, altura); 
			
			f = new Font("Arial", 16, FontStyle.Bold);
			string codBanco = cb.banco.codBanco+"-"+cb.banco.digito;

			w = GetComprimentoStringCentrimentros(f, codBanco);
			rep.RenderDirectText(RetornaStringPosicao(4.7f), RetornaStringPosicao(linha+0.45f), codBanco, RetornaStringPosicao(w), f, Color.Black, AlinharEsquerda());
            
            f = new Font("Arial", 13, FontStyle.Bold);

            string linhaDigitavel;
			//lbNossoNumero.Text = ebb.CalculoNossoNumeroFormatado(gr.numGuia.ToString() + gr.numParcela.ToString(), cb.codCarteira);
            string tipoAmbiente = Parametro.GetParametroPorNome("TipoAmbiente").ToUpper();
            if (tipoAmbiente == "PRODUCAO")
                linhaDigitavel = ebb.CalculoLinhaDigitavel(cb.banco.codBanco, guia.numGuia.ToString() + (guia.numParcela.ToString().PadLeft(2, '0')), cb.codAgencia, cb.numConvenioLider, cb.codCarteira, cb.numConta, guia.dataVencimento, guia.valorGuia);
            else
                linhaDigitavel = "00000.00000 00000.000000 00000.000000 0 00000000000000";// "0".PadLeft(47, '0');
			w = GetComprimentoStringCentrimentros(f, linhaDigitavel);
			rep.RenderDirectText(RetornaStringPosicao(7.1f), RetornaStringPosicao(linha+0.45f), linhaDigitavel,RetornaStringPosicao(w), f, Color.Black, AlinharEsquerda());
			linha = altura;
			
			float yfim = linha + alturaLinha;
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, yfim);
			f = new Font("Arial", 10, FontStyle.Bold);

            da = Parametro.GetParametroPorNome("DescricaoLocalPagamento"); //"Pagável em qualquer banco até a data de vencimento";
			
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Local de Pagamento", da);
			DesenharLinhaVertical(14, linha, yfim);
			f = new Font("Arial", 14, FontStyle.Bold);
			w = GetComprimentoStringCentrimentros(f, guia.dataVencimento.ToString("dd/MM/yyyy"));
			DesenharDescricaoCampo(linhaDireita, linha, "Vencimento", "");
			rep.RenderDirectText(RetornaStringPosicao(15.3f),RetornaStringPosicao(linha), guia.dataVencimento.ToString("dd/MM/yyyy"), RetornaStringPosicao(w),
				f, Color.Black, AlinharCentro());

			linha = yfim;
			yfim = linha + alturaLinha;

			//linha agencia
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, yfim);
			f = new Font("Arial", 10);
			string cedente = Parametro.GetParametroPorNome("TextoCedenteBoleto");
			if (cedente.Trim() == "")
			{
				string nomePM = Parametro.GetParametroPorNome("NomePrefeitura");
				string secretariaPM = Parametro.GetParametroPorNome("NomeSecretaria");
				cedente = nomePM + " - " + secretariaPM;
			}
			linha -= 0.02f;
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Cedente", cedente);
			DesenharLinhaVertical(linhaDireita, linha, yfim);
            string agenciacodigocedente = ebb.CalculoAgenciaCedente(cb.codAgencia, cb.digitoAgencia, cb.numConvenioLider, cb.numConta, cb.digitoConta);//ebb.CalculoAgenciaCedente(cb.codAgencia, cb.numConvenioLider, cb.numConta);
			//agenciacodigocedente = cb.codAgencia.Trim()+"-"+cb.digitoAgencia.Trim()+"/"+cb.numConta.Trim()+"-"+cb.digitoConta.Trim();
			w = GetComprimentoStringCentrimentros(f, agenciacodigocedente);
			DesenharDescricaoCampo(linhaDireita, linha, "Agência Código Cedente", "");
			rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f),RetornaStringPosicao(linha+0.3f), agenciacodigocedente, 
				RetornaStringPosicao(w),	f, Color.Black, AlinharDireita());

			linha = yfim-0.05f;
			yfim = linha + alturaLinha;
			
			//linha nosso numero
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, yfim);
			//linha data
			DesenharLinhaVertical(3.2f, linha, yfim);
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Data Documento", guia.dataEmissao.ToString("dd/MM/yyyy"));
			//linha n doc
			DesenharLinhaVertical(6, linha, yfim);
			DesenharDescricaoCampo(3.2f, linha, "Nº do Documento", guia.numGuia.ToString()+"-"+guia.numParcela.ToString());
			DesenharDescricaoCampo(6, linha, "Espécie Doc.", "");
			//linha especie
			DesenharLinhaVertical(9, linha, yfim);
			DesenharDescricaoCampo(9, linha, "Aceite", "");
			//linha aceite
			DesenharLinhaVertical(10.5f, linha, yfim);
			DesenharDescricaoCampo(10.5f, linha, "Data do Processamento", DateTime.Now.ToString("dd/MM/yyyy"));
			DesenharLinhaVertical(linhaDireita, linha, yfim);
			//nosso numero
			string nossonumero;
			nossonumero = ebb.CalculoNossoNumeroFormatado(guia.numGuia.ToString() + (guia.numParcela.ToString().PadLeft(2,'0')), cb);//cb.numConvenioLider, cb.codCarteira);
			DesenharDescricaoCampo(linhaDireita, linha, "Nosso Número", "");
			w = GetComprimentoStringCentrimentros(f, nossonumero);
			rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f),RetornaStringPosicao(linha+0.3f), nossonumero, 
				RetornaStringPosicao(w), f, Color.Black, AlinharDireita());

			linha = yfim;
			yfim = linha + alturaLinha;

			//linha valor do documento
			DesenharRetagulo(pontoXInicial, linha, pontoXFinal, yfim);
			//linha data
			DesenharLinhaVertical(3.2f, linha, yfim);
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Uso do Banco", "");
			//linha carteira
			DesenharLinhaVertical(4.5f, linha, yfim);
			DesenharDescricaoCampo(3.2f, linha, "Carteira", cb.codCarteira);
			//linha especie
			DesenharLinhaVertical(8.5f, linha, yfim);
			DesenharDescricaoCampo(4.5f, linha, "Espécie", "Real");
			//linha valor
			DesenharLinhaVertical(10.5f, linha, yfim);
			DesenharDescricaoCampo(8.5f, linha, "Quantidade", "");
			DesenharDescricaoCampo(10.5f, linha, "Valor", "");
			DesenharLinhaVertical(linhaDireita, linha, yfim);
			DesenharDescricaoCampo(linhaDireita, linha, "(=) Valor do Documento", "");
			if(guia.valorGuia > 999)
			{
				w = GetComprimentoStringCentrimentros(f, guia.valorGuia.ToString("0,000.00"));
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f),RetornaStringPosicao(linha+0.3f), guia.valorGuia.ToString("0,000.00"), 
					RetornaStringPosicao(w),	f, Color.Black, AlinharDireita());
			}
			else
			{
				w = GetComprimentoStringCentrimentros(f, guia.valorGuia.ToString("0.00"));
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-w-0.2f),RetornaStringPosicao(linha+0.3f), guia.valorGuia.ToString("0.00"), 
					RetornaStringPosicao(w),	f, Color.Black, AlinharDireita());

			}

			linha = yfim;
			yfim = linha + alturaLinha;

			float linhaaux=0;
			//Retangulo desconto
			DesenharRetagulo(linhaDireita, linha, pontoXFinal, linha+alturaLinha);
			DesenharDescricaoCampo(linhaDireita, linha, "(-) Desconto/Abatimento", "");
			linhaaux = linha+alturaLinha;
			//Retangulo  outras deducoes
			DesenharRetagulo(linhaDireita, linhaaux, pontoXFinal, linhaaux+alturaLinha);
			DesenharDescricaoCampo(linhaDireita, linhaaux, "(-) Outras Deduções", "");
			linhaaux += alturaLinha;
			//retangulo mora
			DesenharRetagulo(linhaDireita, linhaaux, pontoXFinal, linhaaux+alturaLinha);
			DesenharDescricaoCampo(linhaDireita, linhaaux, "(+) Mora/Multa", "");
			linhaaux += alturaLinha;
			//retangulo outros acrescimos
			DesenharRetagulo(linhaDireita, linhaaux, pontoXFinal, linhaaux+alturaLinha);
			DesenharDescricaoCampo(linhaDireita, linhaaux, "(+) Outros Acréscimos", "");
			linhaaux += alturaLinha;
			//retangulo outros acrescimos
			DesenharRetagulo(linhaDireita, linhaaux, pontoXFinal, linhaaux+alturaLinha);
			DesenharDescricaoCampo(linhaDireita, linhaaux, "(=) Valor Total Cobrado", "");
			linhaaux += alturaLinha;
			//retangulo observacoes
			DesenharRetagulo(pontoXInicial, linha, linhaDireita, linhaaux);
			DesenharDescricaoCampo(posicaoEsquerda, linha, "Instruções (texto de responsabilidade exclusiva do Cedente)", "");
            Font ft = new Font("Arial", 8);
			string info1 = Parametro.GetParametroPorNome("Info1Guia");
			if(info1 == null)
				info1 = " ";
			else if(info1.Equals(""))
				info1 = " ";
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
				RetornaStringPosicao(linha+0.4f), info1, RetornaStringPosicao(13.5f),
				ft, Color.Black, AlinharEsquerda());
            
			string info2 = Parametro.GetParametroPorNome("Info2Guia");
			if(info2 == null)
				info2 = " ";
			else if(info2.Equals(""))
				info2 = " ";
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
				RetornaStringPosicao(linha+0.93f), info2, RetornaStringPosicao(13.5f),
				ft, Color.Black, AlinharEsquerda());
			string info3 = Parametro.GetParametroPorNome("Info3Guia");
			if(info3 == null)
				info3 = " ";
			else if(info3.Equals(""))
				info3 = " ";
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
				RetornaStringPosicao(linha+1.4f), info3, RetornaStringPosicao(13.5f),
				ft, Color.Black, AlinharEsquerda());
			string info4 = Parametro.GetParametroPorNome("Info4Guia");
			if(info4 == null)
				info4 = " ";
			else if(info4.Equals(""))
				info4 = " ";
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
				RetornaStringPosicao(linha+1.9f), info4, RetornaStringPosicao(13.5f),
				ft, Color.Black, AlinharEsquerda());
			string info5 = guia.info1;
            ft = new Font("Arial", 8,FontStyle.Bold);
			if(info5 != null)
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
					RetornaStringPosicao(linha+2.6f), info5, RetornaStringPosicao(13.5f),
					ft, Color.Black, AlinharEsquerda());

			ft = new Font("Arial", 10, FontStyle.Bold);
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda),
				RetornaStringPosicao(linhaaux-0.5f), "NÃO RECEBER APÓS "+guia.dataValidade.ToString("dd/MM/yyyy"), RetornaStringPosicao(13.5f),
				ft, Color.Black, AlinharEsquerda());
			DesenharRetagulo(pontoXInicial, linhaaux, pontoXFinal, linhaaux+1.1f);
			//retangulo Sacado
			DesenharDescricaoCampo(posicaoEsquerda, linhaaux, "Sacado", "");
            int tamNomeContribuinte = guia.contribuinte.nome.Length;
            
            string contrib = guia.contribuinte.inscricao.Trim() + " - " + guia.contribuinte.nome.Substring(0, tamNomeContribuinte > 60 ? 60 : tamNomeContribuinte).ToUpper() + " - CPF/CNPJ: " + guia.contribuinte.numDocReceita;
            
            w = GetComprimentoStringCentrimentros(f, contrib);
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+0.5f), RetornaStringPosicao(linhaaux+0.3f),
				contrib, RetornaStringPosicao(w), f, Color.Black, AlinharEsquerda());
			h = GetAlturaStringCentrimentros(f, contrib);
			linhaaux += h;
			Contribuinte c = guia.contribuinte;
			c.Retrieve();
			Logradouro l = c.logradouroCarta;
			string linhaenderecoCarta = "";
			if(l != null)
			{
				l.Retrieve();
                if (l.nome != "N/D" && l.nome.ToUpper() != "NÃO DISPONÍVEL")
                {
                    l.cidade.Retrieve();
                    linhaenderecoCarta = "";
                    if (l.tipoLogradouro != null)
                        linhaenderecoCarta += l.tipoLogradouro.descricao.Trim() + " ";
                    if (l.nome != null)
                        linhaenderecoCarta += l.nome.Trim();
                    if (c.numeroCarta != null)
                        linhaenderecoCarta += ", " + c.numeroCarta.Trim() + " ";
                    if (c.complementoCarta != null)
                        linhaenderecoCarta += c.complementoCarta.Trim();
                    if (l.bairro != null)
                        linhaenderecoCarta += " " + l.bairro.nome.Trim();
                    if (l.cidade != null)
                    {
                        linhaenderecoCarta += " " + l.cidade.nome.Trim();
                        if (l.cidade.uf != null)
                            linhaenderecoCarta += "-" + l.cidade.uf.uf.Trim();
                    }
                }
                else
                {
                    linhaenderecoCarta = "";
                }
			}
			w = GetComprimentoStringCentrimentros(f, linhaenderecoCarta);
			rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+0.5f), RetornaStringPosicao(linhaaux+0.1f),
				linhaenderecoCarta, RetornaStringPosicao(w), f, Color.Black, AlinharEsquerda());
			h = GetComprimentoStringCentrimentros(f, linhaenderecoCarta);
			//linhaaux += h;
			da = "AUTENTICAÇÃO MECÂNICA";
			w = GetComprimentoStringCentrimentros(fontDescricao, da);
			rep.RenderDirectText(RetornaStringPosicao(linhaDireita+0.7f), RetornaStringPosicao(linhaaux+0.6f), 
				da, RetornaStringPosicao(w), fontDescricao, Color.Gray, AlinharEsquerda());

            da = "FICHA DE COMPENSAÇÃO";
            w = GetComprimentoStringCentrimentros(fontCompensa, da);
            rep.RenderDirectText(RetornaStringPosicao(linhaDireita + 0.7f), RetornaStringPosicao(linhaaux + 0.9f),
                da, RetornaStringPosicao(w), fontCompensa, Color.Gray, AlinharEsquerda()); 

			//codigo de barras
			string nossonum = guia.numGuia.ToString()+(guia.numParcela.ToString().PadLeft(2, '0'));

            string codBarras;
            if (tipoAmbiente == "PRODUCAO")
                codBarras = ebb.CalculoCodigoBarras(cb.banco.codBanco, nossonum, cb.codAgencia, cb.numConvenioLider, cb.codCarteira, cb.numConta, guia.dataVencimento, guia.valorGuia);
            else
                codBarras = "0".PadLeft(47, '0');
            //string codBarras = ebb.CalculoCodigoBarras(cb.banco.codBanco, nossonum, cb.codAgencia, cb.numConvenioLider, cb.codCarteira, cb.numConta, guia.dataVencimento, guia.valorGuia);

			Bitmap objBitmap = Utils.DesenhaCodigoBarrasInterleaved2of5(codBarras);

			String pathAbsoluto = PathAplicacao + "\\" + Parametro.GetParametroPorNome("SubDiretorioTempFigCodBarra") + "\\"; 
			if(!Directory.Exists(pathAbsoluto))
				Directory.CreateDirectory(pathAbsoluto);
			string pathRelativo = Parametro.GetParametroPorNome("SubDiretorioTempFigCodBarra") + "/";
			//string nomeArquivo = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + gr.oid.ToString() + ".jpg";
			string nomeArquivo = "CodBarra" + guia.oid.ToString() + ".jpg";
			//objBitmap.Save(pathAbsoluto + nomeArquivo, System.Drawing.Imaging.ImageFormat.Jpeg);
			rep.RenderDirectImage(RetornaStringPosicao(posicaoEsquerda+1f), RetornaStringPosicao(linhaaux+0.6f), objBitmap);
            
            objBitmap.Dispose();			
			//ThreadApagarArquivo.ApagarArquivo(pathAbsoluto + nomeArquivo);										
		}

		/// <summary>
		/// Desenha o verso do boleto
		/// </summary>
		private void DesenharVerso()
		{
			if(frenteVerso)
			{
				rep.NewPage();
				float alt = 10f;
				float posesq = 2.0f;
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(alt),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(alt));

				/*Dados da PM*/

				string nomePM = Parametro.GetParametroPorNome("NomePrefeitura");
				string secretariaPM = Parametro.GetParametroPorNome("NomeSecretaria");
				string enderecoCompletoPM = Parametro.GetParametroPorNome("EnderecoCompleto");
				Font f = new Font("Arial", 16, FontStyle.Bold);
				float w = GetComprimentoStringCentrimentros(f, nomePM);
				float h = GetAlturaStringCentrimentros(f, nomePM);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(alt+0.5f), nomePM, 
					w, f, Color.Black, AlinharEsquerda());

				//Gerar imagem selo correio
				if(Parametro.GetParametroPorNome("ArquivoLogotipoCorreio") != "")
				{
					//Image imgArquivoLogotipoCorreio = Image.FromFile(Parametro.GetParametroPorNome("ArquivoLogotipoCorreio"));
					C1.C1PrintDocument.ImageAlignDef ia = new C1.C1PrintDocument.ImageAlignDef();
					//20mm
					rep.RenderDirectImage(RetornaStringPosicao(posicaoEsquerda+15f), RetornaStringPosicao(alt+0.05f), imgArquivoLogotipoCorreio, 3.1f, 3.1f, ia);
					//imgArquivoLogotipoCorreio.Dispose();

					string numcont = "Nº "+Parametro.GetParametroPorNome("NumContratoCorreio");
					f = new Font("Arial", 8);
					w = GetComprimentoStringCentrimentros(f, numcont);
					rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+15.5f), RetornaStringPosicao(alt+1.51f),
						numcont, w, f, Color.Black, AlinharEsquerda());

					string cli = Parametro.GetParametroPorNome("ClienteCorreio");
					f = new Font("Arial", 8);
					w = GetComprimentoStringCentrimentros(f, cli);
					rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+15.5f), RetornaStringPosicao(alt+1.9f),
						cli, w, f, Color.Black, AlinharEsquerda());
				}

				f = new Font("Arial", 15);
				w = GetComprimentoStringCentrimentros(f, secretariaPM);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(alt+0.5f + h),
					secretariaPM, w, f, Color.Black, AlinharEsquerda());
				h += GetAlturaStringCentrimentros(f, secretariaPM);
				f = new Font("Arial", 12);
				w = GetComprimentoStringCentrimentros(f, enderecoCompletoPM);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(alt+0.5f + h),
					enderecoCompletoPM, w, f, Color.Black, AlinharEsquerda());
				h = alt+0.5f + h + 1f;
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(h),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(h));
				h += 0.5f;

				/* Dados do contribuinte */
				f = new Font("Arial", 12, FontStyle.Bold);
				Contribuinte c = guia.contribuinte;
				c.Retrieve();
				string razao = c.nome.ToUpper();
				w = GetComprimentoStringCentrimentros(f, razao);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(h),
					razao, w, f, Color.Black, AlinharEsquerda());
				h += GetAlturaStringCentrimentros(f, razao);
				f = new Font("Arial", 11);
				string nomefant = " ";
				if(c.nomeFantasia != null)
					nomefant = c.nomeFantasia.ToUpper();
				w = GetComprimentoStringCentrimentros(f, nomefant);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(h),
					nomefant, w, f, Color.Black, AlinharEsquerda());
				h += GetAlturaStringCentrimentros(f, nomefant);
				Logradouro l = c.logradouroCarta;
                string endereco = "";
                if (l != null)
                {
                    l.Retrieve();
                    if ((l.nome != "N/D" && l.nome.ToUpper() != "NÃO DISPONÍVEL") || c.enderecoAntigo == null || c.enderecoAntigo.Trim() == "")
                    {
                        if (l.tipoLogradouro != null)
                            endereco = l.tipoLogradouro.descricao + " " + l.nome + ", " + c.numeroCarta + " " + c.complementoCarta;
                        else
                            endereco = l.nome + ", " + c.numeroCarta + " " + c.complementoCarta;
                        w = GetComprimentoStringCentrimentros(f, endereco);
                        rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(h),
                            endereco, w, f, Color.Black, AlinharEsquerda());
                        h += GetAlturaStringCentrimentros(f, endereco);
                        l.cidade.Retrieve();
                        UF u = l.cidade.uf;
                        if (l.bairro != null)
                            endereco = l.bairro.nome + "  " + l.cidade.nome + "-" + u.uf + " CEP: " + l.cep;
                        else
                            endereco = l.cidade.nome + "-" + u.uf + " CEP: " + l.cep;
                    }
                    else
                    {
                        endereco = c.enderecoAntigo;
                    }
                }
                else if(c.enderecoAntigo != null && c.enderecoAntigo != "")
                    endereco = c.enderecoAntigo;

				w = GetComprimentoStringCentrimentros(f, endereco);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(h),
					endereco, w, f, Color.Black, AlinharEsquerda());
				h += GetAlturaStringCentrimentros(f, endereco);
				endereco = "CAE: "+c.inscricao+" -  GUIA: "+guia.numGuia.ToString()+"-"+guia.numParcela.ToString("00");
				w = GetComprimentoStringCentrimentros(f, endereco);
				rep.RenderDirectText(RetornaStringPosicao(posesq), RetornaStringPosicao(h),
					endereco, w, f, Color.Black, AlinharEsquerda());
				h += GetAlturaStringCentrimentros(f, endereco);
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(h),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(h));
				h += 2f;
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(h),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(h));

				h += 3f;

				/* Retangulo dados entrega */

				DesenharRetagulo(posicaoEsquerda+1, h, pontoXFinal-1.5f, h+2.5f);

				//linha vertical meio
				rep.RenderDirectLine(RetornaStringPosicao(posicaoEsquerda+8), RetornaStringPosicao(h),
					RetornaStringPosicao(posicaoEsquerda+8), RetornaStringPosicao(h+2.5f));

				//meia linha vertical
				rep.RenderDirectLine(RetornaStringPosicao(posicaoEsquerda+5f), RetornaStringPosicao(h+1.25f),
					RetornaStringPosicao(posicaoEsquerda+5f), RetornaStringPosicao(h+2.5f));

				rep.RenderDirectLine(RetornaStringPosicao(posicaoEsquerda+1), RetornaStringPosicao(h+1.25f),
					RetornaStringPosicao(posicaoEsquerda+8), RetornaStringPosicao(h+1.25f));
			
				C1.C1PrintDocument.C1DocStyle cs = new C1.C1PrintDocument.C1DocStyle(rep);
				cs.TextAngle = 180;
				cs.TextColor = Color.Black;
				f = new Font("Arial", 8);
				cs.Font = f;
				string s = "Data";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+7.8f-w),RetornaStringPosicao(h+2.2f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				s = "Reintegrado ao Serviço Postal em";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+4.8f-3.5f),RetornaStringPosicao(h+1.9f),
					s,RetornaStringPosicao(3.5f), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)*2), cs); 
			
				s = "Assinatura do Entregador";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda+7.8f-w),RetornaStringPosicao(h+0.9f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				/*cinco retangulos pequenos*/
				DesenharRetagulo(pontoXFinal-2.3f, h+2.1f, pontoXFinal-2, h+2.4f);

				s = "Mudou-se";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-2.4f-w),RetornaStringPosicao(h+2.1f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				DesenharRetagulo(pontoXFinal-2.3f, h+1.6f, pontoXFinal-2, h+1.9f);

				s = "Endereço Insuficiente";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-2.4f-w),RetornaStringPosicao(h+1.6f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				DesenharRetagulo(pontoXFinal-2.3f, h+1.1f, pontoXFinal-2, h+1.4f);

				s = "Não Existe o Nº Indicado";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-2.4f-w),RetornaStringPosicao(h+1.1f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				DesenharRetagulo(pontoXFinal-2.3f, h+0.6f, pontoXFinal-2, h+0.9f);

				s = "Desconhecido";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-2.4f-w),RetornaStringPosicao(h+0.6f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 
                
				DesenharRetagulo(pontoXFinal-2.3f, h+0.1f, pontoXFinal-2, h+0.4f);

				s = "Recusado";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-2.4f-w),RetornaStringPosicao(h+0.1f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 
                
				/*quatro retangulos pequenos */
				DesenharRetagulo(pontoXFinal-7.3f, h+2.1f, pontoXFinal-7, h+2.4f);

				s = "Não Procurado";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-7.4f-w),RetornaStringPosicao(h+2.1f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 
                
				DesenharRetagulo(pontoXFinal-7.3f, h+1.6f, pontoXFinal-7, h+1.9f);

				s = "Ausente";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-7.4f-w),RetornaStringPosicao(h+1.6f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 
                
				DesenharRetagulo(pontoXFinal-7.3f, h+1.1f, pontoXFinal-7, h+1.4f);

				s = "Falecido";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-7.4f-w),RetornaStringPosicao(h+1.1f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				DesenharRetagulo(pontoXFinal-7.3f, h+0.6f, pontoXFinal-7, h+0.9f);

				s = "Inf. Escrita - Porteiro/Síndico";
				w = GetComprimentoStringCentrimentros(f, s);
				rep.RenderDirectText(RetornaStringPosicao(pontoXFinal-7.4f-w),RetornaStringPosicao(h+0.6f),
					s,RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, s)), cs); 

				/* Dados da PM */
				f = new Font("Arial", 12);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;
			
				float w1 = GetComprimentoStringCentrimentros(f, enderecoCompletoPM.Trim());

				f = new Font("Arial", 14);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;
				float w2 = GetComprimentoStringCentrimentros(f, secretariaPM.Trim());
                
				f = new Font("Arial", 15, FontStyle.Bold);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;
				float w3 = GetComprimentoStringCentrimentros(f, nomePM.Trim());

				w = Math.Max(w1, w2);
				w = Math.Max(w, w3);

				f = new Font("Arial", 12);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;

				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda-1.5f), RetornaStringPosicao(h+4.5f), enderecoCompletoPM.Trim(), 
					RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, enderecoCompletoPM)), cs); 

				f = new Font("Arial", 14);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda-1.5f), RetornaStringPosicao(h+5.2f), secretariaPM.Trim(), 
					RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, secretariaPM)), cs); 

				f = new Font("Arial", 15, FontStyle.Bold);
				cs.Font = f;
				cs.TextAlignHorz = C1.C1PrintDocument.AlignHorzEnum.Left;
				rep.RenderDirectText(RetornaStringPosicao(posicaoEsquerda-1.5f), RetornaStringPosicao(h+6f), nomePM.Trim(), 
					RetornaStringPosicao(w), RetornaStringPosicao(GetAlturaStringCentrimentros(f, nomePM)), cs); 

				//linha final
				rep.RenderDirectLine(RetornaStringPosicao(pontoXInicial), RetornaStringPosicao(h+7f),
					RetornaStringPosicao(pontoXFinal), RetornaStringPosicao(h+7f));
			}//fim if frenteVerso
		}
	}
}
