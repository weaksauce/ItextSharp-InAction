using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using G6;
using ObjetosDeNegocio;
using System.Xml;
using System.Xml.XPath;
using System.Collections;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;
using G6.Persistence;


namespace NFeI
{
    public abstract class __Danfe
    {
        protected static CultureInfo ci_br = new CultureInfo("pt-BR");
        protected static CultureInfo ci_us = new CultureInfo("en-US");

        protected enum TipoPagina : int { Frente, Verso, Continuacao };
        protected string[] forms = new string[3];

        protected Nfei nfei;
        protected XmlDocument xdoc;
        protected XmlNode xmlNFE;
        protected bool hasFieldDesconto;

        public static __Danfe GetDanfe(Nfei _nfei)
        {
            _nfei.Retrieve();
            // 1 - Retrato , 2 - Paisagem
            string nomeParametro = _nfei.formatoImpressaoDanfe == 1 ? "FormularioRetrato" : "FormularioPaisagem";
            string dir = Path.Combine(Utils.GetPathAplicacao(), @"forms\");
            if (System.IO.Directory.Exists(dir))
            {
                ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(_nfei.emitente, nomeParametro);
                if (pe != null)
                {
                    string[] formularios = pe.conteudo.Split(';');
                    //string[] files = System.IO.Directory.GetFiles(dir, formularios[0], SearchOption.TopDirectoryOnly);
                    //if (files.Length == 1)
                    //{
                    //    PdfReader r = new PdfReader(files[0]);
                    //    if (r.AcroFields.Fields.Contains(Getkey(r.AcroFields, "codigoBarrasC[0]")))
                    //        return new DanfeV3(_nfei);
                    //    else
                    //        return new NewDanfe(_nfei);
                    //}
                    if (formularios.Length < 1)
                        return new Danfe(_nfei);
                    if (Regex.IsMatch(formularios[0], "V3"))
                        return new DanfeV3(_nfei);
                    if (Regex.IsMatch(formularios[0], "V4"))
                        return new DanfeV4(_nfei);
                    else
                        return new NewDanfe(_nfei);
                }
            }
            return new Danfe(_nfei);
        }

        protected static string Getkey(AcroFields af, string nome)
        {
            string retkey = "";
            foreach (DictionaryEntry de in af.Fields)
            {
                string[] key = de.Key.ToString().Split(".".ToCharArray());
                retkey = key[0] + "." + key[1] + "." + nome;
                break;
            }
            return retkey;
        }

        public abstract string GerarDanfe();

        protected void GetProxPDFStamper(int numPagina, out TipoPagina tipo, out PdfStamper ps, out MemoryStream ms)
        {
            if (numPagina == 1)
                tipo = TipoPagina.Frente;
            else
            {
                if ((numPagina % 2) == 0)
                {
                    tipo = TipoPagina.Verso;
                    if (forms[(int)tipo] == null)
                        tipo = TipoPagina.Continuacao;
                }
                else
                    tipo = TipoPagina.Continuacao;
                if (forms[(int)tipo] == null)
                    tipo = TipoPagina.Frente;
            }

            PdfReader r = new PdfReader(
                new RandomAccessFileOrArray(forms[(int)tipo]), null);
            ms = new MemoryStream();
            ps = new PdfStamper(r,  ms);
        }
        
        protected string GetValorXML(string _path)
        {
            string ret = G6XmlUtils.GetInnerXml(xmlNFE, _path) ?? "";
            return ret;
        }

        protected string format(string valor, int decimais)
        {
            if (valor == "" || valor == null)
                return "";
            Decimal d = Convert.ToDecimal(valor, ci_us);
            string result;
            switch (decimais)
            {
                case 0:
                    result = d.ToString("#,##0", ci_br);
                    break;
                case 2:
                    result = d.ToString("#,##0.00", ci_br);
                    break;
                case 4:
                    result = d.ToString("#,##0.0000", ci_br);
                    break;
                default:
                    result = valor;
                    break;
            }
            return result;
        }

        protected string decode(string val)
        {
            val = val.Replace("&lt;", "<");
            val = val.Replace("&gt;", ">");
            val = val.Replace("&apos;", "'");
            val = val.Replace("&quot;", "\"");
            val = val.Replace("&amp;", "&");
            return val;
        }

        protected void SetMultValue(AcroFields af, string tag, string value)
        {
            for (int i = 0; i < 10; i++)
            {
                string campo = tag + "[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, campo)))
                    af.SetField(campo, value);
                else
                    break;
            }
        }
    }

    public class Danfe : __Danfe
    {
        Document doc;
        PdfWriter writer;
        // MemoryStream msXML;

        float posEsquerda = 28.35f;
        float posInferior = 28.35f;
        float posDireita = 580.65f;
        float posTop = 813.65f;
        float tamanhoBorda = 0.2f;
        float ultlinha = 0f;

        bool indicaMaisdeUmaFolha = false;
        int posicaoItemProdutos = 0;

        

        public Danfe(Nfei _nfei)
        {
            this.nfei = _nfei;
            this.nfei.Retrieve();
            PersistentCriteria pc = new PersistentCriteria("NfeiEArquivoLote");
            pc.AddSelectEqualTo("nfei", this.nfei);
            if (pc.Perform() > 0)
            {
                NfeiEArquivoLote na = (NfeiEArquivoLote)pc[0];
                na.Retrieve();
                ArquivoLote arq = na.arquivoLote;
                arq.Retrieve();
                //XmlTextReader xt = new XmlTextReader(new StringReader(arq.xmlAssinado));
                //xt.EntityHandling = EntityHandling.ExpandEntities;
                                
                xdoc = new XmlDocument();
                //xdoc.Load(xt);
                //xdoc.LoadXml(xt.ReadContentAsString());
                xdoc.LoadXml(arq.xmlAssinado);
                //XmlTextReader xt = new XmlTextReader(new StringReader(arq.xmlAssinado));
                xmlNFE = G6XmlUtils.FindXmlNodeByAttribute(xdoc, "enviNFe/NFe/infNFe/Id", "NFe" + this.nfei.identificador);

            }
        }

        public override string GerarDanfe()
        {
            doc = new Document(PageSize.A4, 28.35f, 28.35f, 28.35f, 28.35f);
            string nome = "nfe"+DateTime.Now.ToString("yyyyMMddhhmmss")+"_"+nfei.numero.ToString()+".pdf";
            string dirDanfe = Parametro.GetParametroPorNome("DiretorioDanfe");
            string pathcompleto = Utils.GetPathAplicacao() + "\\"+dirDanfe+"\\" + nome;
            writer = PdfWriter.GetInstance(doc, new FileStream(pathcompleto, FileMode.Create));
            doc.Open();


            Rectangle rec = new Rectangle(posEsquerda, posTop, posDireita - GetPontosMM(3), posInferior + GetPontosMM(15));
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            
            rec.BorderColor = Color.BLACK;
            //rec.BackgroundColor = Color.CYAN;
   
            doc.Add(rec);

            SetImagemLogotipo(nfei.emitente);

            SetCabecalhoEmitente();

            SetCabecalhoDANFE();

            SetCabecalhoCodigoBarra();

            SetCabecalhoNaturezaNumeroSerie();

            SetCabecalhoInscricaoEstadualCNPJ();

            SetDestinatarioRemetente();

            SetFatura();

            SetCalculoImposto();

            SetTransportadora();

            SetDadosProduto(posicaoItemProdutos);

            SetValorISSQN();
            
            SetDadosAdicionais(nfei.emitente);

            SetRecibo();

            while(indicaMaisdeUmaFolha)
            {
                NovaPagina();
            }

            doc.Close();           
            //ThreadApagarArquivo.ApagarArquivo(pathcompleto, 300);
            return nome;
        }

        private void NovaPagina()
        {
            doc.NewPage();
            SetImagemLogotipo(nfei.emitente);
            SetCabecalhoEmitente();
            SetCabecalhoDANFE();
            SetCabecalhoCodigoBarra();
            SetCabecalhoNaturezaNumeroSerie();
            SetCabecalhoInscricaoEstadualCNPJ();
            SetDadosProduto(posicaoItemProdutos);
        }

        private void SetRecibo()
        {
            DrawLineSerrilhado(posEsquerda, posInferior + GetPontosMM(14), posDireita - GetPontosMM(3), posInferior + GetPontosMM(14));
            string v = "RECEBEMOS DE " + nfei.emitente.razaoSocial.Trim().ToUpper() + " OS PRODUTOS CONSTANTES DA NOTA FISCAL INDICADO AO LADO";
            Rectangle rec = new Rectangle(posEsquerda, posInferior + GetPontosMM(13), posEsquerda + GetPontosMM(160), posInferior);
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            doc.Add(rec);
            rec = new Rectangle(posEsquerda + GetPontosMM(160), posInferior + GetPontosMM(13), posDireita - GetPontosMM(3), posInferior);
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            doc.Add(rec);
            EscreverTexto(posEsquerda + GetPontosMM(2), posInferior + GetPontosMM(11), v, PdfContentByte.ALIGN_LEFT);
            DrawLineSerrilhado(posEsquerda, posInferior + GetPontosMM(10), posDireita - GetPontosMM(3), posInferior + GetPontosMM(10));
            
            EscreverTexto(posEsquerda + GetPontosMM(173), posInferior + GetPontosMM(11), "NF-e", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(168), posInferior + GetPontosMM(7), "Nº "+nfei.numero.ToString("000000000"), PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(1), posInferior + GetPontosMM(7), "Data de Recebimento", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(25), posInferior + GetPontosMM(7), "Identificação e Assinatura do Recebedor", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(172), posInferior + GetPontosMM(4), "Série " + nfei.serie.ToString(), PdfContentByte.ALIGN_LEFT);
            
        }

        private void SetImagemLogotipo(Emitente em)
        {
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "NomeArquivoLogotipo");
            string nomeLogo = pe.conteudo;

            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(Utils.GetPathAplicacao() + "\\imagem\\" + nomeLogo);
            img.ScaleAbsolute(35, 35);
            float px = posEsquerda + GetPontosMM(2);
            float py = posTop - GetPontosMM(13);
            img.SetAbsolutePosition(px, py);
            doc.Add(img);
        }

        private void SetCabecalhoEmitente()
        {
            Rectangle tc = new Rectangle(posEsquerda, posTop, posEsquerda + GetPontosMM(85), posTop - GetPontosMM(16));
            tc.BorderWidth = tamanhoBorda;
            tc.BorderWidthBottom = tamanhoBorda;
            tc.BorderWidthLeft = tamanhoBorda;
            tc.BorderWidthRight = tamanhoBorda;
            tc.BorderWidthTop = tamanhoBorda;
            tc.BorderColor = Color.BLACK;
            doc.Add(tc);
            Emitente em = nfei.emitente;

            em.Retrieve();
            em.logradouro.Retrieve();
            em.logradouro.cidade.Retrieve();
            EscreverTexto(posEsquerda + GetPontosMM(17), posTop - GetPontosMM(5), em.razaoSocial.ToUpper(), PdfContentByte.ALIGN_LEFT);
            string log = em.logradouro.tipoLogradouro.descricao.Trim()+" "+ em.logradouro.nome.Trim() + ", " + em.numeroLogradouro.Trim();
            if (em.complementoLogradouro != null)
                log += " "+em.complementoLogradouro.Trim();
            if (em.logradouro.bairro != null)
                log += " - " + em.logradouro.bairro.nome.Trim() ;
            EscreverTexto(posEsquerda + GetPontosMM(17), posTop - GetPontosMM(8), log.ToUpper(), PdfContentByte.ALIGN_LEFT);
            string cidadeuf = em.logradouro.cidade.nome.Trim() + " - " + em.logradouro.cidade.uf.uf;
            string cep = " CEP " + Utils.Format(em.logradouro.cep, "99.999-999");
            cidadeuf += cep;
            EscreverTexto(posEsquerda + GetPontosMM(17), posTop - GetPontosMM(11), cidadeuf.ToUpper(), PdfContentByte.ALIGN_LEFT);

            string contatos = "";
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "Contatos");
            if (pe != null)
				contatos = pe.conteudo;
			EscreverTexto(posEsquerda + GetPontosMM(17), posTop - GetPontosMM(14), contatos, PdfContentByte.ALIGN_LEFT);

        }

        private void SetCabecalhoCodigoBarra()
        {
            Rectangle tc = new Rectangle(posEsquerda+GetPontosMM(85), posTop, posDireita - GetPontosMM(3), posTop - GetPontosMM(22));
            tc.BorderWidth = tamanhoBorda;
            tc.BorderWidthBottom = tamanhoBorda;
            tc.BorderWidthLeft = tamanhoBorda;
            tc.BorderWidthRight = tamanhoBorda;
            tc.BorderWidthTop = tamanhoBorda;
            tc.BorderColor = Color.WHITE;
            doc.Add(tc);
            
            System.Drawing.Bitmap bmp = BarCode128c.DesenhaCodigoBarrasCode128c(nfei.identificador.Trim(), 47, 1);

            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bmp, System.Drawing.Imaging.ImageFormat.Jpeg);
			img.ScaleAbsoluteWidth(img.ScaledWidth * 0.8f);
			img.SetAbsolutePosition(posEsquerda + GetPontosMM(99), posTop - GetPontosMM(15));
            doc.Add(img);

            ultlinha = posTop - GetPontosMM(16);

            string idform = Utils.Format(nfei.identificador, "99.99.99.99.99.999.999/9999-99-999-999.999.999-999.999-999.9");
            EscreverTextoValor(posEsquerda + GetPontosMM(99),
                    "Chave de Acesso", idform);
            ultlinha = posTop - GetPontosMM(22);
        }

        private void SetCabecalhoDANFE()
        {
            Rectangle tc = new Rectangle(posEsquerda, posTop - GetPontosMM(16), posEsquerda + GetPontosMM(85), posTop - GetPontosMM(22));
            tc.BorderWidth = tamanhoBorda;
            tc.BorderWidthBottom = tamanhoBorda;
            tc.BorderWidthLeft = tamanhoBorda;
            tc.BorderWidthRight = tamanhoBorda;
            tc.BorderWidthTop = tamanhoBorda;
            tc.BorderColor = Color.BLACK;
            doc.Add(tc);
            BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            
            PdfContentByte cb = writer.DirectContent;

            cb.BeginText();
            cb.SetFontAndSize(helvetica, 7);
            
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, 
                "DANFE - Documento Auxiliar de Nota Fiscal Eletrônica",
                posEsquerda + GetPontosMM(11), posTop - GetPontosMM(20), 0);
            cb.EndText();

        }

        private void SetCabecalhoNaturezaNumeroSerie()
        {

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(100)));
            EscreverTextoValor(    posEsquerda + GetPontosMM(2),  "Natureza da Operação",
                GetValorXML("ide/natOp"));

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(100), posEsquerda + GetPontosMM(125)));
            EscreverTexto(posEsquerda + GetPontosMM(102), posTop - GetPontosMM(24), "1-Saída", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(102), posTop - GetPontosMM(27), "2-Entrada", PdfContentByte.ALIGN_LEFT);
            string v = GetValorXML("ide/tpNF");
            if (v.Equals("0"))
                v = "2";
            EscreverTexto(posEsquerda + GetPontosMM(116), posTop - GetPontosMM(25), "[ "+v+" ]", PdfContentByte.ALIGN_LEFT);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(125), posEsquerda + GetPontosMM(150)));

            EscreverTextoValor(posEsquerda + GetPontosMM(127), 
                  "Série",
                  GetValorXML("ide/serie"));


            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita - GetPontosMM(3)));

            v = GetValorXML("ide/nNF");
            v = v.PadLeft(9, '0');
            EscreverTextoValor(posEsquerda + GetPontosMM(152), 
                     "Número",
                    v);
            ultlinha = GetAlturaLinha();
        }

        private void SetCabecalhoInscricaoEstadualCNPJ()
        {

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(45)));

            EscreverTextoValor(posEsquerda + GetPontosMM(2), 
                      "Inscrição Estadual",
                      GetValorXML("emit/IE"));

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(45), posEsquerda + GetPontosMM(100)));

            EscreverTextoValor(posEsquerda + GetPontosMM(47), 
                  "Inscrição Estadual Subst. Tributário",
                  GetValorXML("emit/IEST"));

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(100), posEsquerda + GetPontosMM(150)));

            string v = GetValorXML("emit/CNPJ");
            if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            EscreverTextoValor(posEsquerda + GetPontosMM(102), 
                 "CNPJ",
                v);


            PersistentCriteria pc = new PersistentCriteria("StatusNfei");
            pc.AddSelectEqualTo("nfei", nfei);
            pc.AddSelectEqualTo("StatusNfei[tipoStatusNfei].TipoStatusNfei[acao]", "AUTORIZACAO");
            pc.OrderBy("oid", TipoOrdenamento.Descendente);
            if (pc.Perform() > 0)
            {
                StatusNfei sn = (StatusNfei)pc[0];
                v = sn.numeroProtocolo.ToString();
            }
            else
                v = "";

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita - GetPontosMM(3)));
            EscreverTextoValor(posEsquerda + GetPontosMM(152),
                "Protocolo de Autorização",
                v);

            ultlinha = GetAlturaLinha();
        }

        private void SetDestinatarioRemetente()
        {
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "DESTINATÁRIO/REMETENTE", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(100)));

            EscreverTextoValor(posEsquerda + GetPontosMM(2), 
                  "Nome/Razão Social",
                  GetValorXML("dest/xNome"));


            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(100), posEsquerda + GetPontosMM(145)));

            string v = GetValorXML("dest/CNPJ");
            if (v.Equals(""))
            {
                v = GetValorXML("dest/CPF");
                if (!v.Equals(""))
                    if (v.Length == 11)
                        v = Utils.Format(v, "cpf");
            }
            else if (v.Length == 14)
                v = Utils.FormatCNPJ(v);

            EscreverTextoValor(posEsquerda + GetPontosMM(102), 
                  "CNPJ/CPF",
                v);

           doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita - GetPontosMM(3)));

            v = GetValorXML("ide/dEmi");

            string[] dts = v.Split(new char[] { '-' });
            v = dts[2] + "/" + dts[1] + "/" + dts[0];
               
            EscreverTextoValor(posEsquerda + GetPontosMM(152), 
                  "Data da Emissão",
                v);

            ultlinha = GetAlturaLinha();

            
            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(65)));

            v = GetValorXML("dest/enderDest/xLgr");
            v += ", " + GetValorXML("dest/enderDest/nro");

            EscreverTextoValor(posEsquerda + GetPontosMM(2), 
                 "Endereço",
                v);

           doc.Add(GetRetangulo(posEsquerda + GetPontosMM(65), posEsquerda + GetPontosMM(120)));

            v = GetValorXML("dest/enderDest/xBairro");

            EscreverTextoValor(posEsquerda + GetPontosMM(67), 
                  "Bairro/Distrito",
                v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(120), posEsquerda + GetPontosMM(145)));

            v = GetValorXML("dest/enderDest/CEP");
            v = Utils.Format(v, "99.999-999");

            EscreverTextoValor(posEsquerda + GetPontosMM(122), 
                  "CEP",
                v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita - GetPontosMM(3)));

            v = GetValorXML("ide/dSaiEnt");

            if (v != null && v != "")
            {

                dts = v.Split(new char[] { '-' });
                v = dts[2] + "/" + dts[1] + "/" + dts[0];
                EscreverTextoValor(posEsquerda + GetPontosMM(152),
                      "Data da Saída/Entrada",
                    v);
            }
            else
                EscreverTextoValor(posEsquerda + GetPontosMM(152),
                      "Data da Saída/Entrada",
                    "");

            ultlinha = GetAlturaLinha();

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(55)));

            v = GetValorXML("dest/enderDest/xMun");
            EscreverTextoValor(posEsquerda + GetPontosMM(2), 
                "Município",
               v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(55), posEsquerda + GetPontosMM(65)));

            v = GetValorXML("dest/enderDest/UF");
            EscreverTextoValor(posEsquerda + GetPontosMM(57), 
                 "UF",
               v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(65), posEsquerda + GetPontosMM(120)));

            v = GetValorXML("dest/IE");

            EscreverTextoValor(posEsquerda + GetPontosMM(67), 
                 "Inscrição Estadual",
               v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(120), posEsquerda + GetPontosMM(145)));
            
            v = GetValorXML("dest/enderDest/fone");
            if (!v.Equals(""))
            {
                if (v.Length == 8)
                    v = Utils.Format(v, "9999-9999");
                else if (v.Length == 10)
                    v = Utils.Format(v, "(99)9999-9999");
            }

            EscreverTextoValor(posEsquerda + GetPontosMM(122), 
                  "Fone/Fax",
                v);


            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita - GetPontosMM(3)));

            EscreverTextoValor(posEsquerda + GetPontosMM(152), 
                  "Hora da Saída",
                "");
            ultlinha = GetAlturaLinha();
        }

        private void SetFatura()
        {
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha  - GetPontosMM(3), "FATURA", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);
            float altlinha = GetAlturaLinha() - GetPontosMM(5);
            float inicioLinha = ultlinha - GetPontosMM(4);

            Rectangle tc = new Rectangle(posEsquerda, ultlinha, posEsquerda + GetPontosMM(21), altlinha);
            tc.BorderWidth = tamanhoBorda;
            tc.BorderWidthBottom = tamanhoBorda;
            tc.BorderWidthLeft = tamanhoBorda;
            tc.BorderWidthRight = tamanhoBorda;
            tc.BorderWidthTop = tamanhoBorda;
            tc.BorderColor = Color.BLACK;
            doc.Add(tc);

            EscreverTextoValor(posEsquerda + GetPontosMM(2), 
                  "Número", "");

            float num1 = posEsquerda + GetPontosMM(2);

            Rectangle tc1 = new Rectangle(posEsquerda + GetPontosMM(21), ultlinha, posEsquerda + GetPontosMM(43), altlinha);
            tc1.BorderWidth = tamanhoBorda;
            tc1.BorderWidthBottom = tamanhoBorda;
            tc1.BorderWidthLeft = tamanhoBorda;
            tc1.BorderWidthRight = tamanhoBorda;
            tc1.BorderWidthTop = tamanhoBorda;
            tc1.BorderColor = Color.BLACK;
            doc.Add(tc1);
            EscreverTextoValor(posEsquerda + GetPontosMM(23), 
                 "Vencimento", "");

            float venc1 = posEsquerda + GetPontosMM(23);

            Rectangle tc2 = new Rectangle(posEsquerda + GetPontosMM(43), ultlinha, posEsquerda + GetPontosMM(65), altlinha);
            tc2.BorderWidth = tamanhoBorda;
            tc2.BorderWidthBottom = tamanhoBorda;
            tc2.BorderWidthLeft = tamanhoBorda;
            tc2.BorderWidthRight = tamanhoBorda;
            tc2.BorderWidthTop = tamanhoBorda;
            tc2.BorderColor = Color.BLACK;
            doc.Add(tc2);
            EscreverTextoValor(posEsquerda + GetPontosMM(45), 
                  "Valor", "");

            float vl1 = posEsquerda + GetPontosMM(45);

            Rectangle tc3 = new Rectangle(posEsquerda + GetPontosMM(65), ultlinha, posEsquerda + GetPontosMM(86), altlinha);
            tc3.BorderWidth = tamanhoBorda;
            tc3.BorderWidthBottom = tamanhoBorda;
            tc3.BorderWidthLeft = tamanhoBorda;
            tc3.BorderWidthRight = tamanhoBorda;
            tc3.BorderWidthTop = tamanhoBorda;
            tc3.BorderColor = Color.BLACK;
            doc.Add(tc3);
            EscreverTextoValor(posEsquerda + GetPontosMM(67), 
                  "Número", "");

            float num2 = posEsquerda + GetPontosMM(67);

            Rectangle tc4 = new Rectangle(posEsquerda + GetPontosMM(86), ultlinha, posEsquerda + GetPontosMM(108), altlinha);
            tc4.BorderWidth = tamanhoBorda;
            tc4.BorderWidthBottom = tamanhoBorda;
            tc4.BorderWidthLeft = tamanhoBorda;
            tc4.BorderWidthRight = tamanhoBorda;
            tc4.BorderWidthTop = tamanhoBorda;
            tc4.BorderColor = Color.BLACK;
            doc.Add(tc4);
            EscreverTextoValor(posEsquerda + GetPontosMM(88), 
                  "Vencimento", "");

            float venc2 = posEsquerda + GetPontosMM(88);

            Rectangle tc5 = new Rectangle(posEsquerda + GetPontosMM(108), ultlinha, posEsquerda + GetPontosMM(130), altlinha);
            tc5.BorderWidth = tamanhoBorda;
            tc5.BorderWidthBottom = tamanhoBorda;
            tc5.BorderWidthLeft = tamanhoBorda;
            tc5.BorderWidthRight = tamanhoBorda;
            tc5.BorderWidthTop = tamanhoBorda;
            tc5.BorderColor = Color.BLACK;
            doc.Add(tc5);
            EscreverTextoValor(posEsquerda + GetPontosMM(110), 
                  "Valor", "");

            float vl2 = posEsquerda + GetPontosMM(110);

            Rectangle tc6 = new Rectangle(posEsquerda + GetPontosMM(130), ultlinha, posEsquerda + GetPontosMM(152), altlinha);
            tc6.BorderWidth = tamanhoBorda;
            tc6.BorderWidthBottom = tamanhoBorda;
            tc6.BorderWidthLeft = tamanhoBorda;
            tc6.BorderWidthRight = tamanhoBorda;
            tc6.BorderWidthTop = tamanhoBorda;
            tc6.BorderColor = Color.BLACK;
            doc.Add(tc6);
            EscreverTextoValor(posEsquerda + GetPontosMM(132), 
                  "Número", "");

            float num3 = posEsquerda + GetPontosMM(132);

            Rectangle tc7 = new Rectangle(posEsquerda + GetPontosMM(152), ultlinha, posEsquerda + GetPontosMM(174), altlinha);
            tc7.BorderWidth = tamanhoBorda;
            tc7.BorderWidthBottom = tamanhoBorda;
            tc7.BorderWidthLeft = tamanhoBorda;
            tc7.BorderWidthRight = tamanhoBorda;
            tc7.BorderWidthTop = tamanhoBorda;
            tc7.BorderColor = Color.BLACK;
            doc.Add(tc7);
            EscreverTextoValor(posEsquerda + GetPontosMM(154), 
                  "Vencimento", "");

            float venc3 = posEsquerda + GetPontosMM(154);

            Rectangle tc8 = new Rectangle(posEsquerda + GetPontosMM(174), ultlinha, posDireita - GetPontosMM(3), altlinha);
            tc8.BorderWidth = tamanhoBorda;
            tc8.BorderWidthBottom = tamanhoBorda;
            tc8.BorderWidthLeft = tamanhoBorda;
            tc8.BorderWidthRight = tamanhoBorda;
            tc8.BorderWidthTop = tamanhoBorda;
            tc8.BorderColor = Color.BLACK;
            doc.Add(tc8);
            EscreverTextoValor(posEsquerda + GetPontosMM(176), 
                  "Valor", "");

            float vl3 = posEsquerda + GetPontosMM(176);

            //duplicatas

            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "dup");

            int i = 1;
            for(int k=0; k < ar.Count; k++)
            {
                XmlNode xn = (XmlNode)ar[k];
                string num = G6XmlUtils.GetInnerXml(xn, "nDup");
                string venc = G6XmlUtils.GetInnerXml(xn, "dVenc");
                string vlr = G6XmlUtils.GetInnerXml(xn, "vDup");
                vlr = vlr.Replace('.', ',');
                if (i == 1)
                {
                    EscreverTexto(num1, inicioLinha, num, PdfContentByte.ALIGN_LEFT);
                    string[] dts = venc.Split(new char[] { '-' });
                    string v = dts[2] + "/" + dts[1] + "/" + dts[0];
                    EscreverTexto(venc1, inicioLinha, v, PdfContentByte.ALIGN_LEFT);
                    EscreverTexto(vl1, inicioLinha, vlr, PdfContentByte.ALIGN_LEFT);
                    i++;
                }
                else if (i == 2)
                {
                    EscreverTexto(num2, inicioLinha, num, PdfContentByte.ALIGN_LEFT);
                    string[] dts = venc.Split(new char[] { '-' });
                    string v = dts[2] + "/" + dts[1] + "/" + dts[0];
                    EscreverTexto(venc2, inicioLinha, v, PdfContentByte.ALIGN_LEFT);
                    EscreverTexto(vl2, inicioLinha, vlr, PdfContentByte.ALIGN_LEFT);
                    i++;
                }
                else
                {
                    EscreverTexto(num3, inicioLinha, num, PdfContentByte.ALIGN_LEFT);
                    string[] dts = venc.Split(new char[] { '-' });
                    string v = dts[2] + "/" + dts[1] + "/" + dts[0];
                    EscreverTexto(venc3, inicioLinha, v, PdfContentByte.ALIGN_LEFT);
                    EscreverTexto(vl3, inicioLinha, vlr, PdfContentByte.ALIGN_LEFT);
                    i = 1;
                    inicioLinha -= GetPontosMM(2);
                }
                
                //inicioLinha -= GetPontosMM(2);
            }
            ultlinha = altlinha;

        }

        private void SetCalculoImposto()
        {
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "CÁLCULO DO IMPOSTO", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(46)));
            string v = GetValorXML("total/ICMSTot/vBC");
            v = v.Replace('.', ',');
            EscreverTextoValor( posEsquerda + GetPontosMM(2), "Base de Cálculo do ICMS", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(46), posEsquerda + GetPontosMM(84)));
            v = GetValorXML("total/ICMSTot/vICMS");
            v = v.Replace('.', ',');
            EscreverTextoValor( posEsquerda + GetPontosMM(48), "Valor do ICMS", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(84), posEsquerda + GetPontosMM(122)));
            v = GetValorXML("total/ICMSTot/vBCST"); 
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(86), "Base de Cálculo do ICMS Subst", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(122), posEsquerda + GetPontosMM(160)));
            v = GetValorXML("total/ICMSTot/vST");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(124), "Valor do ICMS Substituição", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(160), posDireita - GetPontosMM(3)));
            v = GetValorXML("total/ICMSTot/vProd");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(162), "Valor Total dos Produtos", v);

            ultlinha = GetAlturaLinha();

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(46)));
            v = GetValorXML("total/ICMSTot/vFrete");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(2), "Valor do Frete", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(46), posEsquerda + GetPontosMM(84)));
            v = GetValorXML("total/ICMSTot/vSeg");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(48), "Valor do Seguro", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(84), posEsquerda + GetPontosMM(122)));
            v = GetValorXML("total/ICMSTot/vOutro");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(86), "Outras Despesas Acessórias", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(122), posEsquerda + GetPontosMM(160)));
            v = GetValorXML("total/ICMSTot/vIPI");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(124), "Valor Total do IPI", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(160), posDireita - GetPontosMM(3)));
            v = GetValorXML("total/ICMSTot/vNF");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(162), "Valor Total da Nota", v);

            ultlinha = GetAlturaLinha();
        }

        private void SetTransportadora()
        {
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "TRANSPORTADOR/VOLUMES TRANSPORTADOS", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(84)));
            string v = GetValorXML("transp/transporta/xNome");
            EscreverTextoValor(posEsquerda + GetPontosMM(2), "Nome/Razão Social", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(84), posEsquerda + GetPontosMM(119)));
            v = "0-Emitente 1-Destinatário   [ "+GetValorXML("transp/modFrete")+" ]";
            EscreverTextoValor(posEsquerda + GetPontosMM(86), "Frete por Conta", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(119), posEsquerda + GetPontosMM(140)));
            v = GetValorXML("transp/veicTransp/placa") ;
            EscreverTextoValor(posEsquerda + GetPontosMM(121), "Placa do Veículo", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(140), posEsquerda + GetPontosMM(148)));
            v = GetValorXML("transp/veicTransp/UF");
            EscreverTextoValor(posEsquerda + GetPontosMM(142), "UF", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(148), posDireita - GetPontosMM(3)));
            v = GetValorXML("transp/transporta/CNPJ");
            v = Utils.FormatCNPJ(v);
            EscreverTextoValor(posEsquerda + GetPontosMM(150), "CNPJ", v);
            
            ultlinha = GetAlturaLinha();

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(84)));
            v = GetValorXML("transp/transporta/xEnder");
            EscreverTextoValor(posEsquerda + GetPontosMM(2), "Endereço", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(84), posEsquerda + GetPontosMM(140)));
            v = GetValorXML("transp/transporta/xMun");
            EscreverTextoValor(posEsquerda + GetPontosMM(86), "Município", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(140), posEsquerda + GetPontosMM(148)));
            v = GetValorXML("transp/transporta/UF");
            EscreverTextoValor(posEsquerda + GetPontosMM(142), "UF", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(148), posDireita - GetPontosMM(3)));
            v = GetValorXML("transp/transporta/IE");
            v = Utils.FormatCNPJ(v);
            EscreverTextoValor(posEsquerda + GetPontosMM(150), "Inscrição Estadual", v);

            ultlinha = GetAlturaLinha();

            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(28)));
            v = GetValorXML("transp/vol/qVol");
            EscreverTextoValor(posEsquerda + GetPontosMM(2), "Quantidade", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(28), posEsquerda + GetPontosMM(65)));
            v = GetValorXML("transp/vol/esp");
            EscreverTextoValor(posEsquerda + GetPontosMM(30), "Espécie", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(65), posEsquerda + GetPontosMM(95)));
            v = GetValorXML("transp/vol/marca");
            EscreverTextoValor(posEsquerda + GetPontosMM(67), "Marca", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(95), posEsquerda + GetPontosMM(120)));
            v = GetValorXML("transp/vol/nVol");
            EscreverTextoValor(posEsquerda + GetPontosMM(97), "Numeração", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(120), posEsquerda + GetPontosMM(155)));
            v = GetValorXML("transp/vol/pesoB");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(122), "Peso Bruto", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(155), posDireita - GetPontosMM(3)));
            v = GetValorXML("transp/vol/pesoL");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(157), "Peso Líqüido", v);

            ultlinha = GetAlturaLinha();
        }

        private void SetDadosProduto(int _posicaoItemProduto)
        {
            int alturaProdutos=1;
            if (!indicaMaisdeUmaFolha)
                alturaProdutos = 55;
            DrawLine(posEsquerda, ultlinha - GetPontosMM(10), posEsquerda, posInferior + GetPontosMM(alturaProdutos));
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "DADOS DO PRODUTO/SERVIÇO", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);

            doc.Add(GetRetangulo2(posEsquerda, posEsquerda + GetPontosMM(12)));
            EscreverTexto(posEsquerda + GetPontosMM(1), ultlinha - GetPontosMM(3), "Cód. Prod.", PdfContentByte.ALIGN_LEFT);
            float pCodProd = posEsquerda + GetPontosMM(1);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(12), posEsquerda + GetPontosMM(60)));
            EscreverTexto(posEsquerda + GetPontosMM(14), ultlinha - GetPontosMM(3), "Descrição do Produto/Serviço", PdfContentByte.ALIGN_LEFT);
            float pDescricao = posEsquerda + GetPontosMM(14);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(60), posEsquerda + GetPontosMM(73)));
            EscreverTexto(posEsquerda + GetPontosMM(61), ultlinha - GetPontosMM(3), "NCM/SH", PdfContentByte.ALIGN_LEFT);
            float pNCM = posEsquerda + GetPontosMM(61);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(73), posEsquerda + GetPontosMM(80)));
            EscreverTexto(posEsquerda + GetPontosMM(75), ultlinha - GetPontosMM(3), "CST", PdfContentByte.ALIGN_LEFT);
            float pCST = posEsquerda + GetPontosMM(75);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(80), posEsquerda + GetPontosMM(90)));
            EscreverTexto(posEsquerda + GetPontosMM(82), ultlinha - GetPontosMM(3), "CFOP", PdfContentByte.ALIGN_LEFT);
            float pCFOP = posEsquerda + GetPontosMM(82);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(90), posEsquerda + GetPontosMM(97)));
            EscreverTexto(posEsquerda + GetPontosMM(91), ultlinha - GetPontosMM(3), "Unid.", PdfContentByte.ALIGN_LEFT);
            float pUnid = posEsquerda + GetPontosMM(91);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(97), posEsquerda + GetPontosMM(112)));
            EscreverTexto(posEsquerda + GetPontosMM(102), ultlinha - GetPontosMM(3), "Quant.", PdfContentByte.ALIGN_LEFT);
            float pQtd = posEsquerda + GetPontosMM(111);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(112), posEsquerda + GetPontosMM(125)));
            EscreverTexto(posEsquerda + GetPontosMM(115), ultlinha - GetPontosMM(3), "V. Unit.", PdfContentByte.ALIGN_LEFT);
            float pVlUnit = posEsquerda + GetPontosMM(124);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(125), posEsquerda + GetPontosMM(140)));
            EscreverTexto(posEsquerda + GetPontosMM(127), ultlinha - GetPontosMM(3), "V. Total", PdfContentByte.ALIGN_LEFT);
            float pVlTot = posEsquerda + GetPontosMM(139);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(140), posEsquerda + GetPontosMM(153)));
            EscreverTexto(posEsquerda + GetPontosMM(142), ultlinha - GetPontosMM(3), "BC ICMS", PdfContentByte.ALIGN_LEFT);
            float pBCICMS = posEsquerda + GetPontosMM(152);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(153), posEsquerda + GetPontosMM(166)));
            EscreverTexto(posEsquerda + GetPontosMM(155), ultlinha - GetPontosMM(3), "V. ICMS", PdfContentByte.ALIGN_LEFT);
            float pICMS = posEsquerda + GetPontosMM(165);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(166), posEsquerda + GetPontosMM(177)));
            EscreverTexto(posEsquerda + GetPontosMM(168), ultlinha - GetPontosMM(3), "V. IPI", PdfContentByte.ALIGN_LEFT);
            float pIPI = posEsquerda + GetPontosMM(176);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(177), posEsquerda + GetPontosMM(184)));
            EscreverTexto(posEsquerda + GetPontosMM(178), ultlinha - GetPontosMM(2), "Alíq", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(178), ultlinha - GetPontosMM(4), "ICMS", PdfContentByte.ALIGN_LEFT);
            float pAliqICMS = posEsquerda + GetPontosMM(183);

            doc.Add(GetRetangulo2(posEsquerda + GetPontosMM(184), posDireita - GetPontosMM(3)));
            EscreverTexto(posEsquerda + GetPontosMM(185), ultlinha - GetPontosMM(2), "Alíq", PdfContentByte.ALIGN_LEFT);
            EscreverTexto(posEsquerda + GetPontosMM(185), ultlinha - GetPontosMM(4), "IPI", PdfContentByte.ALIGN_LEFT);
            float pAliqIPI = posEsquerda + GetPontosMM(188);


            DrawLine(pDescricao - GetPontosMM(2), ultlinha, pDescricao - GetPontosMM(2), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(60), ultlinha, posEsquerda + GetPontosMM(60), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(73), ultlinha, posEsquerda + GetPontosMM(73), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(80), ultlinha, posEsquerda + GetPontosMM(80), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(90), ultlinha, posEsquerda + GetPontosMM(90), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(97), ultlinha, posEsquerda + GetPontosMM(97), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(112), ultlinha, posEsquerda + GetPontosMM(112), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(125), ultlinha, posEsquerda + GetPontosMM(125), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(140), ultlinha, posEsquerda + GetPontosMM(140), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(153), ultlinha, posEsquerda + GetPontosMM(153), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(166), ultlinha, posEsquerda + GetPontosMM(166), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(177), ultlinha, posEsquerda + GetPontosMM(177), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posEsquerda + GetPontosMM(184), ultlinha, posEsquerda + GetPontosMM(184), posInferior + GetPontosMM(alturaProdutos));
            DrawLine(posDireita - GetPontosMM(3), ultlinha, posDireita - GetPontosMM(3), posInferior + GetPontosMM(alturaProdutos));

            DrawLine(posEsquerda, posInferior + GetPontosMM(alturaProdutos), posDireita - GetPontosMM(3), posInferior + GetPontosMM(alturaProdutos));
            ultlinha -= GetPontosMM(6);
            

            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "prod");
            for (int i = _posicaoItemProduto; i < ar.Count; i++)
            {
                 XmlNode xn = (XmlNode)ar[i];
                 string codProd = G6XmlUtils.GetInnerXml(xn, "cProd");
                 string descprod = G6XmlUtils.GetInnerXml(xn, "xProd");
                 descprod = descprod.Replace("||", "\n");
                 XmlNode xinf = G6XmlUtils.FindXmlNode(xn.ParentNode, "infAdProd");
                 if(xinf != null)   
                    descprod = descprod.Trim()+xinf.InnerText ;
                 descprod = descprod.Trim();
                 descprod = descprod.Replace("&lt;", "<");
                 descprod = descprod.Replace("&gt;", ">");
                 descprod = descprod.Replace("&apos;", "'");
                 descprod = descprod.Replace("&quot;", "\"");
                 descprod = descprod.Replace("&amp;", "&");
                 string ncm = G6XmlUtils.GetInnerXml(xn, "NCM");
                 string cfop = G6XmlUtils.GetInnerXml(xn, "CFOP");
                 string uni = G6XmlUtils.GetInnerXml(xn, "uTrib");
                 string qtde = G6XmlUtils.GetInnerXml(xn, "qTrib");
                 if (qtde != null)
                     qtde = qtde.Replace('.', ',');
                 else
                     qtde = "0";
                 decimal fqtde = Convert.ToDecimal(qtde);

                 string vlunit = G6XmlUtils.GetInnerXml(xn, "vUnTrib");
                 if (vlunit != null)
                     vlunit = vlunit.Replace('.', ',');
                 else
                     vlunit = "0";
                 decimal fvlunit = Convert.ToDecimal(vlunit);

                 string vltotal = G6XmlUtils.GetInnerXml(xn, "vProd");
                 if (vltotal != null)
                     vltotal = vltotal.Replace('.', ',');
                 else
                     vltotal = "0";
                 decimal fvltotal = Convert.ToDecimal(vltotal);
                 fvltotal = Utils.RetornaDuasCasasDecimais(fvltotal);
                 vltotal = fvltotal.ToString("0.00");

                 ArrayList ai = G6XmlUtils.GetAllXmlNodes(xmlNFE, "imposto");
                 XmlNode xi = (XmlNode)ai[i];

                 string orig = null;
                 string cst="";
                 string bc="";
                 string icms = "";
                 string ipi = "";
                 string aliqicms = "";
                 string aliqipi = "";
                 
                 orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/orig");
                 if (orig != null)
                 {
                     cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/CST");
                     bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/vBC");
                     icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/vICMS");
                     aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/pICMS");
                 }
                 else
                 {
                     orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/orig");
                     if (orig != null)
                     {
                         cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/CST");
                         bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/vBC");
                         icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/vICMS");
                         aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/pICMS");
                     }
                     else
                     {
                         orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/orig");
                         if (orig != null)
                         {
                             cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/CST");
                             bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/vBC");
                             icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/vICMS");
                             aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/pICMS");
                         }
                         else
                         {
                             orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS30/orig");
                             if (orig != null)
                             {
                                 cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS30/CST");
                             }
                             else
                             {
                                 orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS40/orig");
                                 if (orig != null)
                                 {
                                     cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS40/CST");
                                 }
                                 else
                                 {
                                     orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/orig");
                                     if (orig != null)
                                     {
                                         cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/CST");
                                         bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/vBC");
                                         icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/vICMS");
                                         aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/pICMS");
                                     }
                                     else
                                     {
                                         orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS60/orig");
                                         if (orig != null)
                                         {
                                             cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS60/CST");
                                         }
                                         else
                                         {
                                             orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/orig");
                                             if (orig != null)
                                             {
                                                 cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/CST");
                                                 bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/vBC");
                                                 icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/vICMS");
                                                 aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/pICMS");
                                             }
                                             else
                                             {
                                                 orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/orig");
                                                 if (orig != null)
                                                 {
                                                     cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/CST");
                                                     bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/vBC");
                                                     icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/vICMS");
                                                     aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/pICMS");
                                                 }
                                             }
                                         }
                                     }
                                 }
                             }
                         }
                     }
                 }//fim else if (orig != null)
                 ipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/vIPI");
                 aliqipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/pIPI");
                 if (ipi == null)
                     ipi = "";
                 else
                     ipi = ipi.Replace('.',',');
                 if (aliqipi == null)
                     aliqipi = "";
                 else
                     aliqipi = aliqipi.Replace('.', ',');

                 bc = bc.Replace('.', ',');
                aliqicms = aliqicms.Replace('.',',');

                int tamtexto = 3;
                 if(descprod.Length > 34)
                    tamtexto = (int)Math.Ceiling((decimal)(descprod.Length / 34)) +3;
                 EscreverTexto(pCodProd, ultlinha, codProd, PdfContentByte.ALIGN_LEFT);
                 EscreverValorComQuebraLinha(posEsquerda + GetPontosMM(13), posEsquerda + GetPontosMM(59),
                      ultlinha - GetPontosMM(tamtexto), ultlinha + GetPontosMM(2), 5, descprod);
                 //EscreverTexto(pDescricao, ultlinha, descprod, PdfContentByte.ALIGN_LEFT);
                 EscreverTexto(pNCM, ultlinha, ncm, PdfContentByte.ALIGN_LEFT);
                 EscreverTexto(pCST, ultlinha, cst, PdfContentByte.ALIGN_LEFT);
                 EscreverTexto(pCFOP, ultlinha, cfop, PdfContentByte.ALIGN_LEFT);
                 EscreverTexto(pUnid, ultlinha, uni, PdfContentByte.ALIGN_LEFT);
                 EscreverTexto(pQtd, ultlinha, qtde, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pVlUnit, ultlinha, vlunit, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pVlTot, ultlinha, vltotal, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pBCICMS, ultlinha, bc, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pICMS, ultlinha, icms, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pIPI, ultlinha, ipi, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pAliqICMS, ultlinha, aliqicms, PdfContentByte.ALIGN_RIGHT);
                 EscreverTexto(pAliqIPI, ultlinha, aliqipi, PdfContentByte.ALIGN_RIGHT);

                 ultlinha -= GetPontosMM(tamtexto);
                 //if (ultlinha < (posInferior + GetPontosMM(55))) 2008/07/01 Luiz Cesar
                 if (ultlinha < (posInferior + GetPontosMM(55)) && (i < (ar.Count - 1)))
                 {
                     indicaMaisdeUmaFolha = true;
                     posicaoItemProdutos = i + 1;
                     break;
                 }
                 else
                     indicaMaisdeUmaFolha = false;

                 //ultlinha -= GetPontosMM(3);
            }
            string v = GetValorXML("transp/vol/lacres/nLacre");
            if (v != null && v != "")
            {
                int tamt = 3;
                if (v.Length > 34)
                    tamt = (int)Math.Ceiling((decimal)(v.Length / 34)) + 3;
                EscreverValorComQuebraLinha(posEsquerda + GetPontosMM(13), posEsquerda + GetPontosMM(59),
                ultlinha - GetPontosMM(tamt), ultlinha, 5, "Lacres: "+v);
                ultlinha -= GetPontosMM(tamt);
            }
            if (!indicaMaisdeUmaFolha)
            {
                v = GetValorXML("total/ICMSTot/vDesc");
                if (v != null && v != "")
                {
                    EscreverTextoValor(posEsquerda + GetPontosMM(13), "", "Valor Desconto: " + v);
                }
            }
            ultlinha = posInferior + GetPontosMM(55);

        }

        private void SetValorISSQN()
        {
            EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "CÁLCULO DO ISSQN", PdfContentByte.ALIGN_LEFT);
            ultlinha = ultlinha - GetPontosMM(5);
            doc.Add(GetRetangulo(posEsquerda, posEsquerda + GetPontosMM(40)));
            string v = GetValorXML("emit/IM");
            EscreverTextoValor(posEsquerda + GetPontosMM(2), "Inscrição Municipal", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(40), posEsquerda + GetPontosMM(100)));
            v = GetValorXML("total/ISSQNtot/vServ");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(42), "Valor Total dos Serviços", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(100), posEsquerda + GetPontosMM(150)));
            v = GetValorXML("total/ISSQNtot/vBC");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(102), "Base de Cálculo do ISSQN", v);

            doc.Add(GetRetangulo(posEsquerda + GetPontosMM(150), posDireita  - GetPontosMM(3)));
            v = GetValorXML("total/ISSQNtot/vISS");
            v = v.Replace('.', ',');
            EscreverTextoValor(posEsquerda + GetPontosMM(152), "Valor do ISSQN", v);

            ultlinha = GetAlturaLinha();
        }

        private void SetDadosAdicionais(Emitente em)
        {
            string v = GetValorXML("ide/tpEmis");
            if (v.Trim() == "1")
                EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), "DADOS ADICIONAIS", PdfContentByte.ALIGN_LEFT);
            else
            {
                v = "DADOS ADICIONAIS                                  *** DANFE EM CONTINGÊNCIA, IMPRESSO EM DECORRÊNCIA DE PROBLEMAS TÉCNICOS. ***  ";
                EscreverTexto(posEsquerda + GetPontosMM(2), ultlinha - GetPontosMM(3), v, PdfContentByte.ALIGN_LEFT);
            }
            ultlinha = ultlinha - GetPontosMM(5);
            Rectangle rec = new Rectangle(posEsquerda, ultlinha, posEsquerda + GetPontosMM(130), posInferior + GetPontosMM(15));
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            doc.Add(rec);
            v = GetValorXML("infAdic/infCpl");
            v = v.Replace('\n',' ');
            v = v.Replace('\r', ' ');
            v = v.Replace("||", "\n");
            EscreverTextoValor(posEsquerda+GetPontosMM(2), "Informações Complementares","");
            EscreverValorComQuebraLinha(posEsquerda + GetPontosMM(2), posEsquerda + GetPontosMM(130),
                ultlinha - GetPontosMM(2), posInferior + GetPontosMM(10),5, v);

            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "Mensagem");
            if (pe != null)
				v = pe.conteudo;
            if (v != "")
                EscreverTextoValor(posEsquerda + GetPontosMM(132), "", v);
            rec = new Rectangle(posEsquerda + GetPontosMM(130), ultlinha, posDireita - GetPontosMM(3), posInferior + GetPontosMM(15));
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            v = GetValorXML("infAdic/infAdFisco");
            EscreverTextoValor(posEsquerda + GetPontosMM(132), "Informações Adicionais de Interesse do Fisco", "");
            if (v != null && v != "")
            {
                v = v.Replace('\n', ' ');
                v = v.Replace('\r', ' ');
                v = v.Replace("||", "\n");
                //EscreverTextoValor(posEsquerda + GetPontosMM(132), "Informações Adicionais de Interesse do Fisco", "");
                EscreverValorComQuebraLinha(posEsquerda + GetPontosMM(132), posDireita - GetPontosMM(3),
                    ultlinha - GetPontosMM(2), posInferior + GetPontosMM(10), 5, v);
            }
            doc.Add(rec);
            doc.Add(rec);
        }

        private void EscreverTextoValor(float posX, string descricao, string valor)
        {
            BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            PdfContentByte cb = writer.DirectContent;

            cb.BeginText();
            cb.SetFontAndSize(helvetica, 6);
            cb.SetColorFill(Color.GRAY);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, descricao, posX, GetPosYDescricao(), 0);
            
            cb.SetColorFill(Color.BLACK);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, valor, posX, GetPosYValor(), 0);
            cb.EndText();

        }

        private void EscreverValorComQuebraLinha(float posXInicio, float posXFinal, float posYInicio, float posYFinal, string valor)
        {
            BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font font = new Font(helvetica, 5, Font.NORMAL);

            PdfContentByte cb = writer.DirectContent;

            ColumnText ct = new ColumnText(cb);
            ct.SetSimpleColumn(posXInicio,posYInicio, posXFinal, posYFinal,10, Element.ALIGN_TOP);
            Chunk p = new Chunk(valor, font);
            
            ct.AddText(p);
            ct.Go();

        }

        private void EscreverValorComQuebraLinha(float posXInicio, float posXFinal, float posYInicio, float posYFinal, float entrelinhas, string valor)
        {
            BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font font = new Font(helvetica, 5, Font.NORMAL);

            PdfContentByte cb = writer.DirectContent;

            ColumnText ct = new ColumnText(cb);
            ct.SetSimpleColumn(posXInicio, posYInicio, posXFinal, posYFinal, entrelinhas, Element.ALIGN_BOTTOM);
            Chunk p = new Chunk(valor, font);

            ct.AddText(p);
            ct.Go();

        }

        private void EscreverTexto(float posX, float posY, string valor, int alinhamento)
        {
            BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            PdfContentByte cb = writer.DirectContent;

            cb.BeginText();
            cb.SetFontAndSize(helvetica, 6);
            cb.ShowTextAligned(alinhamento, valor, posX, posY, 0);
            cb.EndText();

        }

        /// <summary>
        /// Desenha uma linha
        /// </summary>
        /// <param name="posXini">Posicao X inicial</param>
        /// <param name="posYini">Posicao Y inicial</param>
        /// <param name="posXfim">Posicao X final</param>
        /// <param name="posYfim">Posicao Y final</param>
        private void DrawLine(float posXini, float posYini,float posXfim, float posYfim)
        {
            PdfContentByte cb = writer.DirectContent;
            
            cb.SetLineWidth(tamanhoBorda);
            cb.MoveTo(posXfim, posYfim);
            cb.LineTo(posXini, posYini);
            cb.Stroke();
        }
        
        private void DrawLineSerrilhado(float posXini, float posYini,float posXfim, float posYfim)
        {
            PdfContentByte cb = writer.DirectContent;
            
            cb.SetLineWidth(0.5f);
            cb.SetLineDash(15f);
            cb.MoveTo(posXfim, posYfim);
            cb.LineTo(posXini, posYini);
            cb.Stroke();
        }

        //private string GetValorXML(string _path)
        //{
        //    //_path = "ide/natOp"
        //    string ret = G6XmlUtils.GetInnerXml(xmlNFE, _path );
        //    if(ret == null)
        //        ret = "";
        //    return ret;
        //}

        private Rectangle GetRetangulo(float posEsquerdaInicial, float posEsquerdaFinal)
        {
            Rectangle rec = new Rectangle(posEsquerdaInicial, ultlinha, posEsquerdaFinal, GetAlturaLinha());
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            return rec;
        }

        private Rectangle GetRetangulo2(float posEsquerdaInicial, float posEsquerdaFinal)
        {
            Rectangle rec = new Rectangle(posEsquerdaInicial, ultlinha, posEsquerdaFinal, GetAlturaLinha()+GetPontosMM(2));
            rec.BorderWidth = tamanhoBorda;
            rec.BorderWidthBottom = tamanhoBorda;
            rec.BorderWidthLeft = tamanhoBorda;
            rec.BorderWidthRight = tamanhoBorda;
            rec.BorderWidthTop = tamanhoBorda;
            rec.BorderColor = Color.BLACK;
            return rec;
        }
        private float GetPosYDescricao()
        {
            return ultlinha - GetPontosMM(2);
        }

        private float GetPosYValor()
        {
            return ultlinha - GetPontosMM(5);
        }

        private float GetAlturaLinha()
        {
            return ultlinha - GetPontosMM(6);
        }

        private float GetPontosMM(int tamanhoMilimitro)
        {
            return tamanhoMilimitro / 0.3527f;
        }
    }

	public class ConcatenarDanfes
	{
		iTextSharp.text.Document outputDocument;
		StreamWriter outputStream;
		iTextSharp.text.pdf.PdfCopy pdfCopy;

		public ConcatenarDanfes(Stream saida)
		{
			outputStream = new StreamWriter(saida);
			outputStream.AutoFlush = true;
		}
		public void Concat(string file)
		{
			iTextSharp.text.pdf.PdfReader inputDocument = new iTextSharp.text.pdf.PdfReader(file);
            Concat(inputDocument);
		}

        public void Concat(PdfReader inputDocument)
        {
            if (outputDocument == null)
            {
                outputDocument = new iTextSharp.text.Document(inputDocument.GetPageSizeWithRotation(1));
                pdfCopy = new iTextSharp.text.pdf.PdfCopy(outputDocument, outputStream.BaseStream);
                outputDocument.Open();
            }
            // Iterate pages
            int count = inputDocument.NumberOfPages;
            for (int idx = 1; idx <= count; idx++)
            {
                // Get the page from the external document...
                // ...and add it to the output document.
                pdfCopy.AddPage(pdfCopy.GetImportedPage(inputDocument, idx));
            }
            pdfCopy.Flush();
            pdfCopy.FreeReader(inputDocument);
        }

		public void Close()
		{
			outputDocument.Close();
		}
	}

    public class myFont
    {
        public Font font {get; set;}
        public float textRise {get; set;}
    }

    // Primeira Versão com Formulários
    public class NewDanfe : __Danfe
    {
        //Document doc;

        Font font_xProd;
        Font font_infCpl;

        public NewDanfe(Nfei _nfei)
        {
            this.nfei = _nfei;
            PersistentCriteria pc = new PersistentCriteria("NfeiEArquivoLote");
            pc.AddSelectEqualTo("nfei", this.nfei);
            if (pc.Perform() > 0)
            {
                NfeiEArquivoLote na = (NfeiEArquivoLote)pc[0];
                na.Retrieve();
                ArquivoLote arq = na.arquivoLote;
                arq.Retrieve();

                xdoc = new XmlDocument();
                xdoc.LoadXml(arq.xmlAssinado);
                xmlNFE = G6XmlUtils.FindXmlNodeByAttribute(xdoc, "enviNFe/NFe/infNFe/Id", "NFe" + this.nfei.identificador);
            }

            // 1 - Retrato , 2 - Paisagem

            string nomeParametro = _nfei.formatoImpressaoDanfe == 1 ? "FormularioRetrato" : "FormularioPaisagem";
            string dir = Path.Combine(Utils.GetPathAplicacao(), @"forms\");
            /*
            string[] files = System.IO.Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly);
            string lista = String.Join(";", files);
            Regex reF = new Regex("(" + tipo + @"_F_.+?\.pdf)", RegexOptions.IgnoreCase);
            Regex reV = new Regex("(" + tipo + @"_V_.+?\.pdf)", RegexOptions.IgnoreCase);
            Regex reC = new Regex("(" + tipo + @"_C_.+?\.pdf)", RegexOptions.IgnoreCase);
            Match mF = reF.Match(lista);
            forms[(int)TipoPagina.Frente] = Path.Combine(Utils.GetPathAplicacao(), @"forms\") + mF.Groups[1].Value;
            Match mV = reV.Match(lista);
            if (mV.Success)
                forms[(int)TipoPagina.Verso] = Path.Combine(Utils.GetPathAplicacao(), @"forms\") + mV.Groups[1].Value;
            Match mC = reC.Match(lista);
            if (mC.Success)
                forms[(int)TipoPagina.Continuacao] = Path.Combine(Utils.GetPathAplicacao(), @"forms\") + mC.Groups[1].Value;
            */
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(_nfei.emitente, nomeParametro);
            string[] tipoForm = pe.conteudo.Split(';');
            for (int i = 0; i < tipoForm.Length; i++)
            {
                if (!tipoForm[i].Equals("Nenhum"))
                    forms[i] = dir + tipoForm[i];
            }

            PdfReader r = new PdfReader(forms[0]);
            font_xProd = getFonte(r.AcroFields.GetField(Getkey(r.AcroFields, "xProd[0]")));
            font_infCpl = getFonte(r.AcroFields.GetField(Getkey(r.AcroFields, "infCpl")));
            r.Close();
        }

        private Font getFonte(string defFonte)
        {
            //Formato : [Helv|Cour];99.9
            if (defFonte == String.Empty)
                defFonte = "helv;6.0";
            string[] fonte_form = defFonte.Split(new char[] { ';' });
            float fontSize = float.Parse(fonte_form[fonte_form.Length - 1], ci_us);
            string fontName = fonte_form[0];
            BaseFont bf;
            if (fontName.ToUpper() == "COUR")
                bf = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            else if (fontName.ToUpper() == "HELV")
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            else // "TMNR"
                bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            return new Font(bf, fontSize, Font.NORMAL);
        }
        /*
                private string Getkey(AcroFields af, string nome)
                {
                    string retkey = "";
                    foreach (DictionaryEntry de in af.Fields)
                    {
                        string[] key = de.Key.ToString().Split(".".ToCharArray());
                        retkey = key[0] + "." + key[1] + "." + nome;
                        break;
                    }
                    return retkey;
                }
        */
        public override string GerarDanfe()
        {

            string nome = "nfe" + DateTime.Now.ToString("yyyyMMddhhmmss") + "_" + nfei.numero.ToString() + "-new.pdf";
            string dirDanfe = Parametro.GetParametroPorNome("DiretorioDanfe");
            string pathcompleto = Utils.GetPathAplicacao() + "\\" + dirDanfe + "\\" + nome;//Stream ms = new MemoryStream();

            ///////////////////////
            ArrayList ar_ps = new ArrayList();
            ArrayList ar_ms = new ArrayList();
            string infFisco;
            string infCpl;
            bool hasMoreInfos = false;
            int proximoProduto = 0;
            TipoPagina tipo;
            bool hasMoreProdutos = false;
            GetDadosAdicionais(out infCpl, out infFisco);
            do
            {
                PdfStamper ps;
                MemoryStream ms;
                GetProxPDFStamper(ar_ps.Count + 1, out tipo, out ps, out ms);
                ar_ps.Add(ps);
                ar_ms.Add(ms);
                switch (tipo)
                {
                    case TipoPagina.Frente:
                        SetImagemLogotipo(nfei.emitente, ps.AcroFields);
                        SetCabecalhoEmitente(ps.AcroFields);
                        SetCabecalhoCodigoBarra(ps.AcroFields);
                        SetCabecalhoNaturezaNumeroSerie(ps.AcroFields);
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        SetDestinatarioRemetente(ps.AcroFields);
                        SetFatura(ps.AcroFields);
                        SetCalculoImposto(ps.AcroFields);
                        SetTransportadora(ps.AcroFields);
                        SetValorISSQN(ps.AcroFields);
                        SetRecibo(ps.AcroFields);
                        break;
                    case TipoPagina.Verso:
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        break;
                    case TipoPagina.Continuacao:
                        SetImagemLogotipo(nfei.emitente, ps.AcroFields);
                        SetCabecalhoCodigoBarra(ps.AcroFields);
                        SetCabecalhoEmitente(ps.AcroFields);
                        SetCabecalhoNaturezaNumeroSerie(ps.AcroFields);
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        break;
                }
                hasMoreInfos = SetDadosAdicionais(ps, ref infCpl, ref infFisco);
                hasMoreProdutos = SetDadosProduto(ps, ref proximoProduto);
            } while (hasMoreProdutos || hasMoreInfos);


            int qtdFolhas = 0;
            foreach (PdfStamper ps in ar_ps)
            {
                if (ps.AcroFields.Fields.Contains(Getkey(ps.AcroFields, "folha[0]")))
                    qtdFolhas++;
            }
            int nFolha = 1;

            int i = 0;
            Document document = new Document();
            PdfSmartCopy copy = new PdfSmartCopy(document, new FileStream(pathcompleto, FileMode.Create));
            document.Open();
            PdfReader reader;

            foreach (PdfStamper ps in ar_ps)
            {
                if (ps.AcroFields.Fields.Contains(Getkey(ps.AcroFields, "folha[0]")))
                {
                    string folha = nFolha.ToString() + " / " + qtdFolhas.ToString();
                    ps.AcroFields.SetField("folha", folha);
                    nFolha++;
                }
                ps.FormFlattening = true;
                ps.Close();
                reader = new PdfReader(((MemoryStream)ar_ms[i]).ToArray());
                copy.AddPage(copy.GetImportedPage(reader, 1));
                i++;
            }
            document.Close();
            return nome;
        }

        private void SetRecibo(AcroFields af)
        {
            string v = "RECEBEMOS DE " + nfei.emitente.razaoSocial.Trim().ToUpper() + " OS PRODUTOS CONSTANTES DA NOTA FISCAL INDICADO AO LADO";
            af.SetField("msgRecebemos", v);
        }


        private void SetImagemLogotipo(Emitente em, AcroFields af)
        {
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "NomeArquivoLogotipo");

            string nomeLogo = pe.conteudo;

            string file = Path.Combine(Utils.GetPathAplicacao(), "imagem\\" + nomeLogo);
            if (File.Exists(file))
            {
                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(Utils.GetPathAplicacao() + "\\imagem\\" + nomeLogo);
                PushbuttonField btf = af.GetNewPushbuttonFromField("imgLogo");
                btf.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
                btf.ProportionalIcon = true;
                btf.Image = img;
                af.ReplacePushbuttonField("imgLogo", btf.Field);
            }

        }

        private void SetCabecalhoEmitente(AcroFields af)
        {
            Emitente em = nfei.emitente;
            em.Retrieve();
            em.RetrieveLogradouroCidadeUF();

            string log = em.logradouro.tipoLogradouro.descricao.Trim() + " " + em.logradouro.nome.Trim() + ", " + em.numeroLogradouro.Trim();
            if (em.complementoLogradouro != null)
                log += " " + em.complementoLogradouro.Trim();
            if (em.logradouro.bairro != null)
                log += "\n" + em.logradouro.bairro.nome.Trim();

            string cidadeuf = em.logradouro.cidade.nome.Trim() + " - " + em.logradouro.cidade.uf.uf;
            string cep = " CEP " + Utils.Format(em.logradouro.cep, "99.999-999");
            cidadeuf += cep;

            string contatos = "";
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "Contatos");
            if (pe != null)
                contatos = pe.conteudo;

            StringBuilder sb = new StringBuilder(em.razaoSocial.ToUpper());
            sb.Append("\n");
            sb.Append(log + "\n");
            sb.Append(cidadeuf + "\n");
            sb.Append(contatos);
            af.SetField("emitente", sb.ToString());
        }

        private void SetCabecalhoCodigoBarra(AcroFields af)
        {

            System.Drawing.Bitmap bmp = BarCode128c.DesenhaCodigoBarrasCode128c(nfei.identificador.Trim(), 47, 1);
            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bmp, System.Drawing.Imaging.ImageFormat.Jpeg);
            img.ScaleAbsoluteWidth(img.ScaledWidth * 0.8f);

            PushbuttonField btf = af.GetNewPushbuttonFromField("codigoBarras");
            btf.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
            btf.ProportionalIcon = false;
            btf.IconFitToBounds = false;
            btf.ScaleIcon = PushbuttonField.SCALE_ICON_ALWAYS;

            btf.Image = img;
            af.ReplacePushbuttonField("codigoBarras", btf.Field);
        }

        private void SetCabecalhoNaturezaNumeroSerie(AcroFields af)
        {
            af.SetField("natOp", GetValorXML("ide/natOp"));

            string v = GetValorXML("ide/tpNF");
            if (v.Equals("0"))
                v = "2";

            af.SetField("tpNF", v);

            v = GetValorXML("ide/nNF");
            uint u = Convert.ToUInt32(v);
            //v = v.PadLeft(9, '0');
            v = u.ToString("000,000,000", ci_br);

            for (int i = 0; i < 5; i++)
            {
                string nf = "nNF[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, nf)))
                    af.SetField(nf, v);
                else
                    break;
            }

            v = GetValorXML("ide/serie");
            for (int i = 0; i < 5; i++)
            {
                string serie = "serie[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, serie)))
                    af.SetField(serie, v);
                else
                    break;
            }
        }

        private void SetCabecalhoInscricaoEstadualCNPJ(AcroFields af)
        {
            af.SetField("eIE", GetValorXML("emit/IE"));
            af.SetField("eIEST", GetValorXML("emit/IEST"));

            string v = GetValorXML("emit/CNPJ");
            if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            af.SetField("eCNPJ", v);

            if (af.Fields.Contains(Getkey(af, "Protocolo[0]")))
            {
                PersistentCriteria pc = new PersistentCriteria("StatusNfei");
                pc.AddSelectEqualTo("nfei", nfei);
                pc.AddSelectEqualTo("StatusNfei[tipoStatusNfei].TipoStatusNfei[acao]", "AUTORIZACAO");
                pc.OrderBy("oid", TipoOrdenamento.Descendente);
                if (pc.Perform() > 0)
                {
                    StatusNfei sn = (StatusNfei)pc[0];
                    v = sn.numeroProtocolo.ToString();
                }
                else
                    v = "";
                af.SetField("Protocolo", v);
            }


            string idform = Utils.Format(nfei.identificador, "99.99.99.99.99.999.999/9999-99-999-999.999.999-999.999-999.9");
            af.SetField("cNF", idform);
        }

        private void SetDestinatarioRemetente(AcroFields af)
        {
            af.SetField("dxNome", GetValorXML("dest/xNome"));

            string v = GetValorXML("dest/CNPJ");
            if (v.Equals(""))
            {
                v = GetValorXML("dest/CPF");
                if (!v.Equals(""))
                    if (v.Length == 11)
                        v = Utils.Format(v, "cpf");
            }
            else if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            af.SetField("dCnpj", v);

            v = GetValorXML("ide/dEmi");

            string[] dts = v.Split(new char[] { '-' });
            v = dts[2] + "/" + dts[1] + "/" + dts[0];
            af.SetField("dEmi", v);
            v = GetValorXML("dest/enderDest/xLgr");
            v += ", " + GetValorXML("dest/enderDest/nro");
            af.SetField("dxLgr", v);

            v = GetValorXML("dest/enderDest/xBairro");
            af.SetField("dxBairro", v);

            v = GetValorXML("dest/enderDest/CEP");
            v = Utils.Format(v, "99.999-999");
            af.SetField("dCEP", v);

            v = GetValorXML("ide/dSaiEnt");

            if (v != null && v != "")
            {

                dts = v.Split(new char[] { '-' });
                v = dts[2] + "/" + dts[1] + "/" + dts[0];
                af.SetField("dSaiEnt", v);
            }

            v = GetValorXML("dest/enderDest/xMun");
            af.SetField("dxMun", v);

            v = GetValorXML("dest/enderDest/UF");
            af.SetField("dUF", v);

            v = GetValorXML("dest/IE");
            af.SetField("dIE", v);

            v = GetValorXML("dest/enderDest/fone");
            if (!v.Equals(""))
            {
                if (v.Length == 8)
                    v = Utils.Format(v, "9999-9999");
                else if (v.Length == 10)
                    v = Utils.Format(v, "(99)9999-9999");
            }

            af.SetField("dfone", v);
        }

        private void SetFatura(AcroFields af)
        {
            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "dup");
            if (ar.Count > 0)
            {
                for (int k = 0; k < ar.Count; k++)
                {
                    string xdup = "nDup[" + k.ToString() + "]";
                    if (af.Fields.Contains(Getkey(af, xdup)))
                    {
                        string xvenc = "dVenc[" + k.ToString() + "]";
                        string xvlr = "vDup[" + k.ToString() + "]";
                        XmlNode xn = (XmlNode)ar[k];
                        string num = G6XmlUtils.GetInnerXml(xn, "nDup");
                        string venc = G6XmlUtils.GetInnerXml(xn, "dVenc");
                        string[] dts = venc.Split(new char[] { '-' });
                        string v = dts[2] + "/" + dts[1] + "/" + dts[0];
                        string vlr = G6XmlUtils.GetInnerXml(xn, "vDup");
                        vlr = format(vlr, 2);

                        af.SetField(xdup, num);
                        af.SetField(xvenc, v);
                        af.SetField(xvlr, vlr);
                    }
                }
            }
            else
            {
                string fat = GetValorXML("cobr/fat/nFat");
                if (!(fat == "" || fat == null))
                {
                    string vfat = format(GetValorXML("cobr/fat/vLiq"), 2);
                    af.SetField("nDup[0]", fat);
                    af.SetField("vDup[0]", vfat);
                }
            }
        }

        private void SetCalculoImposto(AcroFields af)
        {
            string v = GetValorXML("total/ICMSTot/vBC");
            af.SetField("vBcTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vICMS");
            af.SetField("vICMSTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vBCST");
            af.SetField("vBCSTTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vST");
            af.SetField("vSTTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vProd");
            af.SetField("vProdTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vFrete");
            af.SetField("vFreteTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vSeg");
            af.SetField("vSegTot", format(v, 2));

            if (af.Fields.Contains(Getkey(af, "vDescTot[0]")))
            {
                v = GetValorXML("total/ICMSTot/vDesc");
                af.SetField("vDescTot", format(v, 2));
                hasFieldDesconto = true;
            }

            v = GetValorXML("total/ICMSTot/vOutro");
            af.SetField("vOutroTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vIPI");
            af.SetField("vIPITot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vNF");
            af.SetField("vNFTot", format(v, 2));
        }

        private void SetTransportadora(AcroFields af)
        {
            string v = GetValorXML("transp/transporta/xNome");
            af.SetField("txNome", v);

            v = GetValorXML("transp/modFrete");
            af.SetField("modFrete", v);

            v = GetValorXML("transp/veicTransp/placa");
            af.SetField("tPlacaVeiculo", v);

            v = GetValorXML("transp/veicTransp/UF");
            af.SetField("tvUF", v);

            v = GetValorXML("transp/veicTransp/RNTC");
            af.SetField("rntc", v);

            v = GetValorXML("transp/transporta/CNPJ");
            v = Utils.FormatCNPJ(v);
            af.SetField("tCnpj", v);

            v = GetValorXML("transp/transporta/xEnder");
            af.SetField("txLgr", v);

            v = GetValorXML("transp/transporta/xMun");
            af.SetField("txMun", v);

            v = GetValorXML("transp/transporta/UF");
            af.SetField("tUF", v);

            v = GetValorXML("transp/transporta/IE");
            v = Utils.FormatCNPJ(v);
            af.SetField("tIE", v);

            v = GetValorXML("transp/vol/qVol");
            af.SetField("qVol", v);

            v = GetValorXML("transp/vol/esp");
            af.SetField("esp", v);

            v = GetValorXML("transp/vol/marca");
            af.SetField("marca", v);

            v = GetValorXML("transp/vol/nVol");
            af.SetField("nVol", v);

            v = GetValorXML("transp/vol/pesoB");
            v = v.Replace('.', ',');
            af.SetField("pesoB", v);

            v = GetValorXML("transp/vol/pesoL");
            v = v.Replace('.', ',');
            af.SetField("pesoL", v);
        }

        private bool SetDadosProduto(PdfStamper ps, ref int proximoProduto)
        {
            //BaseFont cour = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //font = new Font(courier, fontSize, Font.NORMAL);

            AcroFields af = ps.AcroFields;
            int qtdLinhas = 0;
            int nroLinhasDescpro;
            for (int i = 0; i < 200; i++)
            {
                string test = "cProd[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, test)))
                    continue;
                qtdLinhas = i;
                break;
            }

            int curLinha = 0;

            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "prod");
            for (int i = proximoProduto; i < ar.Count; i++, proximoProduto++)
            {
                if (curLinha >= qtdLinhas)
                    return true;  // Acabou espaço nesta página, pegue outra...

                XmlNode xn = (XmlNode)ar[i];
                string codProd = G6XmlUtils.GetInnerXml(xn, "cProd");
                string descprod = G6XmlUtils.GetInnerXml(xn, "xProd");
                XmlNode xinf = G6XmlUtils.FindXmlNode(xn.ParentNode, "infAdProd");
                if (xinf != null)
                    descprod = descprod.Trim() + "||" + xinf.InnerText;
                descprod = descprod.Replace("||", "\n");
                descprod = descprod.Trim();
                descprod = descprod.Replace("&lt;", "<");
                descprod = descprod.Replace("&gt;", ">");
                descprod = descprod.Replace("&apos;", "'");
                descprod = descprod.Replace("&quot;", "\"");
                descprod = descprod.Replace("&amp;", "&");

                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, descprod, ps);
                if (nroLinhasDescpro == 0)
                    return true; // tem mais produtos (não coube na pagina)

                string ncm = G6XmlUtils.GetInnerXml(xn, "NCM");
                string cfop = G6XmlUtils.GetInnerXml(xn, "CFOP");
                string uni = G6XmlUtils.GetInnerXml(xn, "uTrib");
                string qtde = G6XmlUtils.GetInnerXml(xn, "qTrib");
                if (qtde == null)
                    qtde = "0";
                qtde = format(qtde, 4);

                string vlunit = G6XmlUtils.GetInnerXml(xn, "vUnTrib");
                if (vlunit == null)
                    vlunit = "0";
                vlunit = format(vlunit, 4);

                string vltotal = G6XmlUtils.GetInnerXml(xn, "vProd");
                if (vltotal == null)
                    vltotal = "0";
                vltotal = format(vltotal, 2);

                ArrayList ai = G6XmlUtils.GetAllXmlNodes(xmlNFE, "imposto");
                XmlNode xi = (XmlNode)ai[i];

                string orig = null;
                string cst = "";
                string bc = "";
                string icms = "";
                string ipi = "";
                string aliqicms = "";
                string aliqipi = "";

                orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/orig");
                if (orig != null)
                {
                    cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/CST");
                    bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/vBC");
                    icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/vICMS");
                    aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS00/pICMS");
                }
                else
                {
                    orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/orig");
                    if (orig != null)
                    {
                        cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/CST");
                        bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/vBC");
                        icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/vICMS");
                        aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS10/pICMS");
                    }
                    else
                    {
                        orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/orig");
                        if (orig != null)
                        {
                            cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/CST");
                            bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/vBC");
                            icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/vICMS");
                            aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS20/pICMS");
                        }
                        else
                        {
                            orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS30/orig");
                            if (orig != null)
                            {
                                cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS30/CST");
                            }
                            else
                            {
                                orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS40/orig");
                                if (orig != null)
                                {
                                    cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS40/CST");
                                }
                                else
                                {
                                    orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/orig");
                                    if (orig != null)
                                    {
                                        cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/CST");
                                        bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/vBC");
                                        icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/vICMS");
                                        aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS51/pICMS");
                                    }
                                    else
                                    {
                                        orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS60/orig");
                                        if (orig != null)
                                        {
                                            cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS60/CST");
                                        }
                                        else
                                        {
                                            orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/orig");
                                            if (orig != null)
                                            {
                                                cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/CST");
                                                bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/vBC");
                                                icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/vICMS");
                                                aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS70/pICMS");
                                            }
                                            else
                                            {
                                                orig = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/orig");
                                                if (orig != null)
                                                {
                                                    cst = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/CST");
                                                    bc = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/vBC");
                                                    icms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/vICMS");
                                                    aliqicms = G6XmlUtils.GetInnerXml(xi, "ICMS/ICMS90/pICMS");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                ipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/vIPI");
                aliqipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/pIPI");
                if (ipi == null)
                    ipi = "";
                else
                    ipi = format(ipi, 2);
                if (aliqipi == null)
                    aliqipi = "";
                else
                    aliqipi = format(aliqipi, 2);

                bc = format(bc, 2);
                aliqicms = format(aliqicms, 2);
                icms = format(icms, 2);

                string vlinha = "[" + curLinha.ToString() + "]";

                af.SetField("cProd" + vlinha, codProd);
                af.SetField("NCM" + vlinha, ncm);
                af.SetField("CST" + vlinha, orig + cst);
                af.SetField("CFOP" + vlinha, cfop);
                af.SetField("uCom" + vlinha, uni);
                af.SetField("qCom" + vlinha, qtde);
                af.SetField("vUnCom" + vlinha, vlunit);
                af.SetField("vProd" + vlinha, vltotal);
                af.SetField("vBCICMS" + vlinha, bc);
                af.SetField("vICMS" + vlinha, icms);
                af.SetField("vIPI" + vlinha, ipi);
                af.SetField("pICMS" + vlinha, aliqicms);
                af.SetField("pIPI" + vlinha, aliqipi);

                curLinha = curLinha + nroLinhasDescpro;
            }

            ArrayList arL = G6XmlUtils.GetAllXmlNodes(xmlNFE, "lacres");
            string lacres = "";
            for (int i = 0; i < arL.Count; i++)
            {
                XmlNode xn = (XmlNode)arL[i];
                string lacre = G6XmlUtils.GetInnerXml(xn, "nLacre");
                lacres = lacres + lacre + ", ";
            }
            //string lacres = GetValorXML("transp/vol/lacres/nLacre");
            if (!(lacres == null || lacres == ""))
                lacres = "\n" + "Lacres: " + lacres;

            string vdesc = "";
            if (!hasFieldDesconto)
            {
                vdesc = GetValorXML("total/ICMSTot/vDesc");
                if (!(vdesc == null || vdesc == "" || vdesc == "0.00"))
                    vdesc = "\n" + "Valor Desconto: " + format(vdesc, 2);
                else
                    vdesc = "";
            }
            string txtAdic = vdesc + lacres;
            if (txtAdic != "")
            {
                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, txtAdic, ps);
                if (nroLinhasDescpro == 0)
                    return true; // dados adic não coube na pagina
            }

            return false; // hasMoreProdutos = false
        }

        private void SetValorISSQN(AcroFields af)
        {
            string v = GetValorXML("emit/IM");
            af.SetField("IM", v);

            v = GetValorXML("total/ISSQNtot/vServ");
            v = v.Replace('.', ',');
            af.SetField("ISSQNtot", v);

            v = GetValorXML("total/ISSQNtot/vBC");
            v = v.Replace('.', ',');
            af.SetField("vBC", v);

            v = GetValorXML("total/ISSQNtot/vISS");
            v = v.Replace('.', ',');
            af.SetField("vISSQN", v);
        }

        private bool SetDadosAdicionais(PdfStamper ps, ref string infCpl, ref string infFisco)
        {
            string tpEmis = GetValorXML("ide/tpEmis");
            if (tpEmis.Trim() != "1")
                ps.AcroFields.SetField("contingencia", "*** DANFE EM CONTINGÊNCIA, IMPRESSO EM DECORRÊNCIA DE PROBLEMAS TÉCNICOS. ***");

            // "xxxx xxxx xxxx /// xxxx xxxx xxxx /// xxxxx xxxxxx xxxxxxx /// xxxxxxxxxxxxxxxxxxx xxxxxxxxxxx /// xxxxxx"
            // "xxxx xxxx xxxx\nxxxx xxxx xxxx\nxxxxx xxxxxx xxxxxxx\nxxxxxxxxxxxxxxxxxxx xxxxxxxxxxx\nxxxxxx"

            if (infCpl != String.Empty)
                infCpl = escreveDadosAdicionais(ps, "infCpl", infCpl);
            if (infFisco != String.Empty)
                infFisco = escreveDadosAdicionais(ps, "infAdFisco", infFisco);
            if (infFisco == String.Empty && infCpl == String.Empty)
                return false; // Não tem mais o que escrever.
            return true;
        }

        private string escreveDadosAdicionais(PdfStamper ps, string campo, string conteudo)
        {
            float[] pos = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, campo));
            PdfContentByte cb = ps.GetOverContent(1);  //writer.DirectContent;
            string infor = conteudo;
            while (true)
            {
                ColumnText ct = new ColumnText(cb);
                ct.SetSimpleColumn(pos[1] + 2, pos[2], pos[3] - 2, pos[4]);
                //ct.AdjustFirstLine = false;
                ct.Leading = font_infCpl.GetCalculatedLeading(1.1f);
                //ct.Alignment = Element.ALIGN_TOP;
                ct.Alignment = Element.ALIGN_BOTTOM;
                //ct.Alignment = Element.ALIGN_BASELINE;
                float Ypos = ct.YLine;

                Chunk p = new Chunk(infor, font_infCpl);
                //ColumnText ctx = ct;
                ct.AddText(p);
                if (!ColumnText.HasMoreText(ct.Go(true)))
                {
                    ct.YLine = Ypos;
                    ct.AddText(p);
                    ct.Go(false);
                    break;
                }
                int ultPos = infor.LastIndexOfAny(new char[] { ' ', '\n' });
                infor = infor.Remove(ultPos).Trim();
                //ct.ClearChunks();
                //ct.YLine = Ypos;
                //ct.FilledWidth = 0.0f;
            }
            //if (infor.Length < conteudo.Length)
            return conteudo.Remove(0, infor.Length).Trim();
        }

        private void GetDadosAdicionais(out string infCpl, out string infFisco)
        {
            string v = GetValorXML("infAdic/infAdFisco");
            infFisco = v.Replace("||", "\n");

            v = GetValorXML("infAdic/infCpl");

            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(nfei.emitente, "Mensagem");
            string mensagem = "";
            if (pe != null)
                mensagem = pe.conteudo;
            if (!String.IsNullOrEmpty(mensagem))
                v = mensagem + "||" + v;
            infCpl = v.Replace("||", "\n");

        }

        //private void EscreverTextoValor(float posX, string descricao, string valor)
        //{
        //    BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

        //    PdfContentByte cb = writer.DirectContent;
        //    cb.BeginText();
        //    cb.SetFontAndSize(helvetica, 6);
        //    cb.SetColorFill(Color.GRAY);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, descricao, posX, GetPosYDescricao(), 0);
        //    cb.SetColorFill(Color.BLACK);
        //    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, valor, posX, GetPosYValor(), 0);
        //    cb.EndText();
        //}

        private int EscreverDescricao(int curlinha, int qtdlinhas, string descricao, PdfStamper ps)
        {
            float[] posFinal = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, "xProd[" + curlinha.ToString() + "]"));
            float[] posInicial = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, "xProd[" + (qtdlinhas - 1).ToString() + "]"));
            //float fieldPage = fieldPositions[0];
            //float fieldLlx = fieldPositions[1];
            //float fieldLly = fieldPositions[2];
            //float fieldUrx = fieldPositions[3];
            //float fieldUry = fieldPositions[4];
            PdfContentByte cb = ps.GetOverContent(1);  //writer.DirectContent;

            ColumnText ct = new ColumnText(cb);
            ct.SetSimpleColumn(posInicial[1], posInicial[2], posFinal[3], posFinal[4]);//, 10, Element.ALIGN_TOP);
            //ct.AdjustFirstLine = false;
            ct.Leading = font_xProd.GetCalculatedLeading(1.1f);
            //ct.Alignment = Element.ALIGN_TOP;
            ct.Alignment = Element.ALIGN_BOTTOM;
            //ct.Alignment = Element.ALIGN_BASELINE;
            Chunk p = new Chunk(descricao, font_xProd);
            float Ypos = ct.YLine;
            ct.AddText(p);

            if (!ColumnText.HasMoreText(ct.Go(true)))
            {
                ct.YLine = Ypos;
                ct.AddText(p);
                ct.Go(false);
                return ct.LinesWritten;
            }
            return 0;

        }

        //private void EscreverValorComQuebraLinha(float posXInicio, float posXFinal, float posYInicio, float posYFinal, float entrelinhas, string valor)
        //{
        //    BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        //    Font font = new Font(helvetica, 5, Font.NORMAL);

        //    PdfContentByte cb = writer.DirectContent;

        //    ColumnText ct = new ColumnText(cb);
        //    ct.SetSimpleColumn(posXInicio, posYInicio, posXFinal, posYFinal, entrelinhas, Element.ALIGN_BOTTOM);
        //    Chunk p = new Chunk(valor, font);

        //    ct.AddText(p);
        //    ct.Go();
        //}
    }

    // DANFE V3 
    public class DanfeV3 : __Danfe
    {
        //Document doc;
		protected ArquivoLote arq;
        //PdfWriter writer;
        protected myFont font_xProd;
        protected myFont font_infCpl;

        public DanfeV3(Nfei _nfei)
        {
            this.nfei = _nfei;
            // Via ArquivoLote
            PersistentCriteria pc = new PersistentCriteria("ArquivoLote");
            pc.AddSelectEqualTo("ArquivoLote[nfeiArquivoLote].NfeiEArquivoLote[nfei]", this.nfei);
            pc.Perform();
            arq = (ArquivoLote)pc[0];
            xdoc = new XmlDocument();
            xdoc.LoadXml(arq.xmlAssinado);
            xmlNFE = G6XmlUtils.FindXmlNodeByAttribute(xdoc, "enviNFe/NFe/infNFe/Id", "NFe" + this.nfei.identificador);

            // 1 - Retrato , 2 - Paisagem

            string nomeParametro = _nfei.formatoImpressaoDanfe == 1 ? "FormularioRetrato" : "FormularioPaisagem";
            string dir = Path.Combine(Utils.GetPathAplicacao(), @"forms\");

            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(_nfei.emitente, nomeParametro);
            string[] tipoForm = pe.conteudo.Split(';');
            for (int i = 0; i < tipoForm.Length; i++)
            {
                if (!tipoForm[i].Equals("Nenhum"))
                    forms[i] = dir + tipoForm[i];
            }

            PdfReader r = new PdfReader(forms[0]);
            font_xProd = getFonte2(r.AcroFields.GetField(Getkey(r.AcroFields, "xProd[0]")));
            font_infCpl = getFonte2(r.AcroFields.GetField(Getkey(r.AcroFields, "infCpl")));
        }

        public override string GerarDanfe()
        {
            string nome = String.Format("nfe_{0}_{1}_{2}-new.pdf", nfei.emitente.numDocReceita, DateTime.Now.ToString("yyyyMMddhhmmss"), nfei.numero.ToString());
            string dirDanfe = Parametro.GetParametroPorNome("DiretorioDanfe");
            string pathcompleto = Utils.GetPathAplicacao() + "\\" + dirDanfe + "\\" + nome;

            ///////////////////////
            ArrayList ar_ps = new ArrayList();
            ArrayList ar_ms = new ArrayList();
            string infCpl;
            bool hasMoreInfos = true;
            int proximoProduto = 0;
            TipoPagina tipo;
            bool hasMoreProdutos = true;
            infCpl = GetDadosAdicionais();
            do
            {
                PdfStamper ps;
                MemoryStream ms;
                GetProxPDFStamper(ar_ps.Count + 1, out tipo, out ps, out ms);
                ar_ps.Add(ps);
                ar_ms.Add(ms);
                switch (tipo)
                {
                    case TipoPagina.Frente:
                        SetImagemLogotipo(ps.AcroFields);
                        SetCabecalhoEmitente(ps.AcroFields);
                        SetCabecalhoCodigoBarra(ps.AcroFields);
                        SetCabecalhoNaturezaNumeroSerie(ps.AcroFields);
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        SetDestinatarioRemetente(ps.AcroFields);
                        SetFatura(ps.AcroFields);
                        SetCalculoImposto(ps.AcroFields);
                        SetTransportadora(ps.AcroFields);                                            
                        SetValorISSQN(ps.AcroFields);
                        SetRecibo(ps.AcroFields);
                        break;
                    case TipoPagina.Verso:
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        break;
                    case TipoPagina.Continuacao:
                        SetImagemLogotipo(ps.AcroFields);
                        SetCabecalhoCodigoBarra(ps.AcroFields);
                        SetCabecalhoEmitente(ps.AcroFields);
                        SetCabecalhoNaturezaNumeroSerie(ps.AcroFields);
                        SetCabecalhoInscricaoEstadualCNPJ(ps.AcroFields);
                        break;
                }
				if (hasMoreInfos)
					hasMoreInfos = SetDadosAdicionais(ps, ref infCpl);
                if (hasMoreProdutos)
					hasMoreProdutos = SetDadosProduto(ps, ref proximoProduto);

                if (ps.AcroFields.Fields.Contains(Getkey(ps.AcroFields, "parceiro[0]")))
                {                     
                    string parceiro = Parametro.GetParametroPorNome("Parceiro");
                    if (parceiro == null)
                        parceiro = "";
                    ps.AcroFields.SetField("parceiro",parceiro);
                }

            } while (hasMoreProdutos || hasMoreInfos);


            int qtdFolhas = 0;
            foreach (PdfStamper ps in ar_ps)
            {
                if (ps.AcroFields.Fields.Contains(Getkey(ps.AcroFields, "folha[0]")))
                    qtdFolhas++;
            }
            int nFolha = 1;

            int i = 0;
            Document document = new Document();
            PdfSmartCopy copy = new PdfSmartCopy(document, new FileStream(pathcompleto, FileMode.Create));
            document.Open();
            PdfReader reader;
            PdfGState gState = new PdfGState();
            gState.FillOpacity = 0.3f;
            gState.StrokeOpacity = 0.3f;

            foreach (PdfStamper ps in ar_ps)
            {
                if (ps.AcroFields.Fields.Contains(Getkey(ps.AcroFields, "folha[0]")))
                {
                    string folha = nFolha.ToString() + " / " + qtdFolhas.ToString();
                    ps.AcroFields.SetField("folha", folha);
                    nFolha++;
                }
                // Se nota de homologação - adiciona marca d'agua
                string v = GetValorXML("ide/tpAmb");
                if (v == "2")
                {
                    PdfContentByte pc = ps.GetOverContent(1);
                    pc.SaveState();
                    pc.SetGState(gState);
                    pc.BeginText();
                    pc.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 60.0f);
                    pc.SetTextRenderingMode(PdfContentByte.TEXT_RENDER_MODE_STROKE);
                    pc.ShowTextAligned(Element.ALIGN_CENTER,
                        "SEM VALOR FISCAL",
                        PageSize.A4.Width / 2,
                        PageSize.A4.Height / 2, 45f);
                    pc.EndText();
                    pc.RestoreState();
                }
                ps.FormFlattening = true;
                ps.Close();
                reader = new PdfReader(((MemoryStream)ar_ms[i]).ToArray());
                copy.AddPage(copy.GetImportedPage(reader, 1));
                i++;
            }
            document.Close();
            return nome;
        }

        protected virtual void SetRecibo(AcroFields af)
        {
            string v = "RECEBEMOS DE " + nfei.emitente.razaoSocial.Trim().ToUpper() + " - " + nfei.emitente.nomeFantasia.Trim() + " OS PRODUTOS CONSTANTES DA NOTA FISCAL INDICADO AO LADO";
            SetMultValue(af, "msgRecebemos", v);
            SetMultValue(af, "exNome", nfei.emitente.razaoSocial.Trim().ToUpper());
        }

        protected virtual void SetImagemLogotipo(AcroFields af)
        {
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(nfei.emitente, "NomeArquivoLogotipo");
            if (pe != null)
            {
                string nomeLogo = pe.conteudo;
                string file = Path.Combine(Utils.GetPathAplicacao(), "imagem\\" + nomeLogo);
                if (File.Exists(file))
                {
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(Utils.GetPathAplicacao() + "\\imagem\\" + nomeLogo);
                    PushbuttonField btf = af.GetNewPushbuttonFromField("imgLogo");
                    btf.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
                    btf.ProportionalIcon = true;
                    btf.Image = img;
                    af.ReplacePushbuttonField("imgLogo", btf.Field);
                }
            }
        }

        protected virtual void SetCabecalhoEmitente(AcroFields af)
        {
            Emitente em = nfei.emitente;
            em.Retrieve();
            em.RetrieveLogradouroCidadeUF();

            string log = em.logradouro.tipoLogradouro.descricao.Trim() + " " + em.logradouro.nome.Trim() + ", " + em.numeroLogradouro.Trim();
            if (em.complementoLogradouro != null)
                log += " " + em.complementoLogradouro.Trim();
            if (em.logradouro.bairro != null)
                log += "\n" + em.logradouro.bairro.nome.Trim();

            string cidadeuf = em.logradouro.cidade.nome.Trim() + " - " + em.logradouro.cidade.uf.uf;
            string cep = " CEP " + Utils.Format(em.logradouro.cep, "99.999-999");
            cidadeuf += cep;

            string contatos = "";
            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(em, "Contatos");
            if (pe != null)
                contatos = pe.conteudo;

            StringBuilder sb = new StringBuilder(em.razaoSocial.ToUpper());
            sb.Append("\n");
            sb.Append(log + "\n");
            sb.Append(cidadeuf + "\n");
            sb.Append(contatos);
            af.SetField("emitente", sb.ToString());
        }

        protected virtual void SetCabecalhoCodigoBarra(AcroFields af)
        {

            System.Drawing.Bitmap bmp = BarCode128c.DesenhaCodigoBarrasCode128c(nfei.identificador.Trim(), 47, 1);
            iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bmp, System.Drawing.Imaging.ImageFormat.Jpeg);
            img.ScaleAbsoluteWidth(img.ScaledWidth * 0.8f);

            PushbuttonField btf = af.GetNewPushbuttonFromField("codigoBarras");
            btf.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
            btf.ProportionalIcon = false;
            btf.IconFitToBounds = false;
            btf.ScaleIcon = PushbuttonField.SCALE_ICON_ALWAYS;

            btf.Image = img;
            af.ReplacePushbuttonField("codigoBarras", btf.Field);
        }

        protected virtual void SetCabecalhoNaturezaNumeroSerie(AcroFields af)
        {
            af.SetField("natOp", GetValorXML("ide/natOp"));

            string v = GetValorXML("ide/tpNF");
            //if (v.Equals("0"))
            //    v = "2";
            af.SetField("tpNF", v);

            v = GetValorXML("ide/nNF");
            uint u = Convert.ToUInt32(v);
            //v = v.PadLeft(9, '0');
            v = u.ToString("000,000,000", ci_br);
            //for (int i = 0; i < 5; i++)
            //{
            //    string nf = "nNF[" + i.ToString() + "]";
            //    if (af.Fields.Contains(Getkey(af, nf)))
            //        af.SetField(nf, v);
            //    else
            //        break;
            //}
            SetMultValue(af, "nNF", v);

            v = GetValorXML("ide/serie");
            //for (int i = 0; i < 5; i++)
            //{
            //    string serie = "serie[" + i.ToString() + "]";
            //    if (af.Fields.Contains(Getkey(af, serie)))
            //        af.SetField(serie, v);
            //    else
            //        break;
            //}
            SetMultValue(af, "serie", v);
        }

        protected virtual void SetCabecalhoInscricaoEstadualCNPJ(AcroFields af)
        {
            af.SetField("eIE", GetValorXML("emit/IE"));
            af.SetField("eIEST", GetValorXML("emit/IEST"));

            string v = GetValorXML("emit/CNPJ");
            if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            af.SetField("eCNPJ", v);

            if (af.Fields.Contains(Getkey(af, "Protocolo[0]")))
            {
                string tpEmis = GetValorXML("ide/tpEmis");
                // se SCAN imprime o protocolo e não o Código de Barras, se DPEC imprime o Recibo
                if (tpEmis == "2" || tpEmis == "5") // FS ou FS-DA
                {
                    string uf = GetValorXML("dest/enderDest/cMun").Substring(0, 2).PadLeft(2, '0');
                    v = GetValorXML("dest/CNPJ");
                    if (v.Equals(""))
                        v = GetValorXML("dest/CPF");
                    string cnpj = v.PadLeft(14, '0');
                    string vNF = Utils.FormatMoney(nfei.valorNotaFiscal).Replace(".", "").Replace(",", "").PadLeft(14, '0');

                    string ICMSp = "2";
                    string ICMSs = "2";
                    if ((float.Parse(GetValorXML("total/ICMSTot/vICMS"))) > 0.00)
                        ICMSp = "1";
                    if ((float.Parse(GetValorXML("total/ICMSTot/vST"))) > 0.00)
                        ICMSs = "1";

                    string dia = nfei.dataEmissao.Day.ToString().PadLeft(2, '0');
                    string dadosNfe = uf + tpEmis + cnpj + vNF + ICMSp + ICMSs + dia;
                    string DV = Utils.Modulo11(dadosNfe);
                    string vCB = dadosNfe + DV;
                    v = Utils.Format(dadosNfe + DV, "9999 9999 9999 9999 9999 9999 9999 9999 9999");

                    af.SetField("labelProtocolo", "DADOS DA NF-e");

                    //CÓDIGO DE BARRAS - Contingência
                    System.Drawing.Bitmap bmp = BarCode128c.DesenhaCodigoBarrasCode128c(vCB.Trim(), 47, 1);
                    iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(bmp, System.Drawing.Imaging.ImageFormat.Jpeg);
                    img.ScaleAbsoluteWidth(img.ScaledWidth * 0.8f);
                    PushbuttonField btf = af.GetNewPushbuttonFromField("codigoBarrasC");
                    btf.Layout = PushbuttonField.LAYOUT_ICON_ONLY;
                    btf.ProportionalIcon = false;
                    btf.IconFitToBounds = false;
                    btf.ScaleIcon = PushbuttonField.SCALE_ICON_ALWAYS;
                    btf.Image = img;
                    af.ReplacePushbuttonField("codigoBarrasC", btf.Field);
                }
                else if (tpEmis == "4")
                {
                    af.SetField("labelProtocolo", "Número de Registro DPEC");
                    
                    PersistentCriteria pc = new PersistentCriteria("DPEC");
                    pc.AddSelectEqualTo("arquivoLote", arq);
                    if (pc.Perform() > 0)
                    {
                        DPEC dpec = (DPEC)pc[0];
                        v = dpec.nRegDPEC.Trim() + " - " + dpec.dhRegDPEC.ToString("dd/MM/yyyy HH:mm:ss");
                    }
                }
                else
                {
                    PersistentCriteria pc = new PersistentCriteria("StatusNfei");
                    pc.AddSelectEqualTo("nfei", nfei);
                    pc.AddSelectEqualTo("StatusNfei[tipoStatusNfei].TipoStatusNfei[acao]", "AUTORIZACAO");
                    pc.OrderBy("oid", TipoOrdenamento.Descendente);
                    if (pc.Perform() > 0)
                    {
                        StatusNfei sn = (StatusNfei)pc[0];
                        v = sn.numeroProtocolo.ToString() + " - " + sn.dataRecebimento.ToString("dd/MM/yyyy HH:mm:ss");
                        af.SetField("labelProtocolo", "PROTOCOLO DE AUTORIZAÇÃO DE USO");
                    }
                }
                af.SetField("Protocolo", v);
            }
            string idform = Utils.Format(nfei.identificador, "9999 9999 9999 9999 9999 9999 9999 9999 9999 9999 9999");
            af.SetField("cNF", idform);
        }

        protected virtual void SetDestinatarioRemetente(AcroFields af)
        {
            //af.SetField("dxNome", decode(GetValorXML("dest/xNome")));
            SetMultValue(af, "dxNome", decode(GetValorXML("dest/xNome")));

            string v = GetValorXML("dest/CNPJ");
            if (v.Equals(""))
            {
                v = GetValorXML("dest/CPF");
                if (!v.Equals(""))
                    if (v.Length == 11)
                        v = Utils.Format(v, "cpf");
            }
            else if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            af.SetField("dCnpj", v);

            v = GetValorXML("ide/dEmi");
            string[] dts = v.Split(new char[] { '-' });
            v = dts[2] + "/" + dts[1] + "/" + dts[0];
            //for (int i = 0; i < 5; i++)
            //{
            //    string campo = "dEmi[" + i.ToString() + "]";
            //    if (af.Fields.Contains(Getkey(af, campo)))
            //        af.SetField(campo, v);
            //    else
            //        break;
            //}
            SetMultValue(af, "dEmi", v);

            v = GetValorXML("dest/enderDest/xLgr");
            v += ", " + GetValorXML("dest/enderDest/nro");
			string xCpl = GetValorXML("dest/enderDest/xCpl");
			if (xCpl.Length > 0)
				v += " /" + xCpl; 
            af.SetField("dxLgr", v);

            v = GetValorXML("dest/enderDest/xBairro");
            af.SetField("dxBairro", v);

            v = GetValorXML("dest/enderDest/CEP");
            v = Utils.Format(v, "99.999-999");
            af.SetField("dCEP", v);

            v = GetValorXML("ide/dSaiEnt");

            if (v != null && v != "")
            {

                dts = v.Split(new char[] { '-' });
                v = dts[2] + "/" + dts[1] + "/" + dts[0];
                af.SetField("dSaiEnt", v);
            }

            //Atribui HORA/SAÍDA
            XmlNode xmlObsCont = G6XmlUtils.FindXmlNodeByAttribute(xmlNFE, "infAdic/obsCont/xCampo", "HoraSaida");
            if (xmlObsCont!=null)
                if (!String.IsNullOrEmpty(xmlObsCont.InnerText))
                    af.SetField("horaSaida", xmlObsCont.InnerText);

            v = GetValorXML("dest/enderDest/xMun");
            af.SetField("dxMun", v);

            v = GetValorXML("dest/enderDest/UF");
            af.SetField("dUF", v);

            v = GetValorXML("dest/IE");
            af.SetField("dIE", v);

            v = GetValorXML("dest/enderDest/fone");
            if (!v.Equals(""))
            {
                if (v.Length == 8)
                    v = Utils.Format(v, "9999-9999");
                else if (v.Length == 10)
                    v = Utils.Format(v, "(99)9999-9999");
            }

            af.SetField("dfone", v);
        }

        protected virtual void SetFatura(AcroFields af)
        {
            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "dup");
            if (ar.Count > 0)
            {
                for (int k = 0; k < ar.Count; k++)
                {
                    string xdup = "nDup[" + k.ToString() + "]";
                    if (af.Fields.Contains(Getkey(af, xdup)))
                    {
                        string xvenc = "dVenc[" + k.ToString() + "]";
                        string xvlr = "vDup[" + k.ToString() + "]";
                        XmlNode xn = (XmlNode)ar[k];
                        string num = G6XmlUtils.GetInnerXml(xn, "nDup");
                        string venc = G6XmlUtils.GetInnerXml(xn, "dVenc");
                        string v = "";
                        if (venc != null)
                        {
                            string[] dts = venc.Split(new char[] { '-' });
                            v = dts[2] + "/" + dts[1] + "/" + dts[0];
                        }
                        string vlr = G6XmlUtils.GetInnerXml(xn, "vDup");

                        if (num == null)
                            num = "";                        
                        if (vlr == null)
                            vlr = "";
                        else 
                            vlr = format(vlr, 2);

                        af.SetField(xdup, num);
                        af.SetField(xvenc, v);
                        af.SetField(xvlr, vlr);
                    }
                }
            }
            else
            {
                if (af.Fields.Contains(Getkey(af, "nDup[0]")))
                {
                    string fat = GetValorXML("cobr/fat/nFat");
                    if (! String.IsNullOrEmpty(fat))
                    {
                        string vfat = format(GetValorXML("cobr/fat/vLiq"), 2);

                        af.SetField("nDup[0]", fat);
                        af.SetField("vDup[0]", vfat);
                    }
                }
            }
        }

        protected virtual void SetCalculoImposto(AcroFields af)
        {
            string v = GetValorXML("total/ICMSTot/vBC");
            af.SetField("vBcTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vICMS");
            af.SetField("vICMSTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vBCST");
            af.SetField("vBCSTTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vST");
            af.SetField("vSTTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vProd");
            af.SetField("vProdTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vFrete");
            af.SetField("vFreteTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vSeg");
            af.SetField("vSegTot", format(v, 2));

            if (af.Fields.Contains(Getkey(af, "vDescTot[0]")))
            {
                v = GetValorXML("total/ICMSTot/vDesc");
                af.SetField("vDescTot", format(v, 2));
                hasFieldDesconto = true;
            }

            v = GetValorXML("total/ICMSTot/vOutro");
            af.SetField("vOutroTot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vIPI");
            af.SetField("vIPITot", format(v, 2));

            v = GetValorXML("total/ICMSTot/vNF");
            af.SetField("vNFTot", format(v, 2));
        }

        protected virtual void SetTransportadora(AcroFields af)
        {
            string v = decode(GetValorXML("transp/transporta/xNome"));
            //af.SetField("txNome", v);
            SetMultValue(af, "txNome", v);

            v = GetValorXML("transp/modFrete");
            af.SetField("modFrete", v);

            v = GetValorXML("transp/veicTransp/placa");
            af.SetField("tPlacaVeiculo", v);

            v = GetValorXML("transp/veicTransp/UF");
            af.SetField("tvUF", v);

            v = GetValorXML("transp/veicTransp/RNTC");
            af.SetField("rntc", v);

            v = GetValorXML("transp/transporta/CNPJ");
            v = Utils.FormatCNPJ(v);
            af.SetField("tCnpj", v);

            v = decode(GetValorXML("transp/transporta/xEnder"));
            af.SetField("txLgr", v);

            v = GetValorXML("transp/transporta/xMun");
            af.SetField("txMun", v);

            v = GetValorXML("transp/transporta/UF");
            af.SetField("tUF", v);

            v = GetValorXML("transp/transporta/IE");
            af.SetField("tIE", v);

            v = GetValorXML("transp/vol/qVol");
            af.SetField("qVol", v);

            v = decode(GetValorXML("transp/vol/esp"));
            af.SetField("esp", v);

            v = decode(GetValorXML("transp/vol/marca"));
            af.SetField("marca", v);

            v = GetValorXML("transp/vol/nVol");
            af.SetField("nVol", v);

            v = GetValorXML("transp/vol/pesoB");
            v = v.Replace('.', ',');
            af.SetField("pesoB", v);

            v = GetValorXML("transp/vol/pesoL");
            v = v.Replace('.', ',');
            af.SetField("pesoL", v);
        }

        protected virtual bool SetDadosProduto(PdfStamper ps, ref int proximoProduto)
        {
            //BaseFont cour = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            //font = new Font(courier, fontSize, Font.NORMAL);

            AcroFields af = ps.AcroFields;
            int qtdLinhas = 0;
            int nroLinhasDescpro;
            for (int i = 0; i < 200; i++)
            {
                string test = "cProd[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, test)))
                    continue;
                qtdLinhas = i;
                break;
            }

            int curLinha = 0;
            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "prod");

            for (int i = proximoProduto; i < ar.Count; i++, proximoProduto++)
            {
                if (curLinha >= qtdLinhas)
                    return true;  // Acabou espaço nesta página, pegue outra...

                XmlNode xn = (XmlNode)ar[i];
                string codProd = G6XmlUtils.GetInnerXml(xn, "cProd");
                string descprod = G6XmlUtils.GetInnerXml(xn, "xProd");
                XmlNode xinf = G6XmlUtils.FindXmlNode(xn.ParentNode, "infAdProd");
                if (xinf != null)
                    descprod = descprod.Trim() + "||" + xinf.InnerText;
                descprod = descprod.Replace("||", "\n");
                descprod = descprod.Trim();
				descprod = decode(descprod);

                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, descprod, ps);
                if (nroLinhasDescpro == 0)
                    return true; // tem mais produtos (não coube na pagina)

                string ncm = G6XmlUtils.GetInnerXml(xn, "NCM");
                if (String.IsNullOrEmpty(ncm))
                    ncm = G6XmlUtils.GetInnerXml(xn, "genero");
                string cfop = G6XmlUtils.GetInnerXml(xn, "CFOP");
                string uni = G6XmlUtils.GetInnerXml(xn, "uCom");
                string qtde = G6XmlUtils.GetInnerXml(xn, "qCom");
                if (qtde == null)
                    qtde = "0";
                qtde = format(qtde, 4);

                string vlunit = G6XmlUtils.GetInnerXml(xn, "vUnCom");
                if (vlunit == null)
                    vlunit = "0";
                vlunit = format(vlunit, 4);

                string vltotal = G6XmlUtils.GetInnerXml(xn, "vProd");
                if (vltotal == null)
                    vltotal = "0";
                vltotal = format(vltotal, 2);

                ArrayList ai = G6XmlUtils.GetAllXmlNodes(xmlNFE, "imposto");
                XmlNode xi = (XmlNode)ai[i];

                string orig = null;
                string cst = "";
                string bc = "";
                string icms = "";
                string bcst = "";
                string icmsst = "";
                string ipi = "";
                string aliqicms = "";
                string aliqipi = "";

				XmlNode nICMS = G6XmlUtils.FindXmlNode(xi, "ICMS");
				if (nICMS != null)
				{
					XmlNode tICMS = nICMS.FirstChild;
					orig = G6XmlUtils.GetInnerXml(tICMS, "orig");
					cst = G6XmlUtils.GetInnerXml(tICMS, "CST");
					bc = G6XmlUtils.GetInnerXml(tICMS, "vBC") ?? "";
					icms = G6XmlUtils.GetInnerXml(tICMS, "vICMS") ?? "";
					aliqicms = G6XmlUtils.GetInnerXml(tICMS, "pICMS") ?? "";
					bcst = G6XmlUtils.GetInnerXml(tICMS, "vBCST") ?? "";
					icmsst = G6XmlUtils.GetInnerXml(tICMS, "vICMSST") ?? "";
				}

                if (af.Fields.Contains(Getkey(af, "vIPI[0]")))
                {
                    ipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/vIPI");
                    if (ipi == null)
                        ipi = "";
                    else
                        ipi = format(ipi, 2);
                }
                if (af.Fields.Contains(Getkey(af, "pIPI[0]")))
                {
                    aliqipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/pIPI");
                    if (aliqipi == null)
                        aliqipi = "";
                    else
                        aliqipi = format(aliqipi, 2);
                }

                bc = format(bc, 2);
                aliqicms = format(aliqicms, 2);
                icms = format(icms, 2);
                bcst = format(bcst, 2);
                icmsst = format(icmsst, 2);

                string vlinha = "[" + curLinha.ToString() + "]";

                af.SetField("cProd" + vlinha, codProd);
                af.SetField("NCM" + vlinha, ncm);
                af.SetField("CST" + vlinha, orig + cst);
                af.SetField("CFOP" + vlinha, cfop);
                af.SetField("uCom" + vlinha, uni);
                af.SetField("qCom" + vlinha, qtde);
                af.SetField("vUnCom" + vlinha, vlunit);
                af.SetField("vProd" + vlinha, vltotal);
                af.SetField("vBCICMS" + vlinha, bc);
                af.SetField("vICMS" + vlinha, icms);
                //SUBST. TRIB
                if (af.Fields.Contains(Getkey(af, "vBCST[0]")))
                    af.SetField("vBCST" + vlinha, bcst);
                if (af.Fields.Contains(Getkey(af, "vICMSST[0]")))
                    af.SetField("vICMSST" + vlinha, icmsst);
                //IPI Opcional
                if (af.Fields.Contains(Getkey(af, "vIPI[0]")))
                    af.SetField("vIPI" + vlinha, ipi);
                af.SetField("pICMS" + vlinha, aliqicms);
                if (af.Fields.Contains(Getkey(af, "pIPI[0]")))
                    af.SetField("pIPI" + vlinha, aliqipi);

                curLinha = curLinha + nroLinhasDescpro;
            }

            ArrayList arL = G6XmlUtils.GetAllXmlNodes(xmlNFE, "lacres");
            string lacres = "";
            for (int i = 0; i < arL.Count; i++)
            {
                XmlNode xn = (XmlNode)arL[i];
                string lacre = G6XmlUtils.GetInnerXml(xn, "nLacre");
                lacres = lacres + lacre + ", ";
            }
            //string lacres = GetValorXML("transp/vol/lacres/nLacre");
            if (!(lacres == null || lacres == ""))
                lacres = "\n" + "Lacres: " + lacres;

            string vdesc = "";
            if (!hasFieldDesconto)
            {
                vdesc = GetValorXML("total/ICMSTot/vDesc");
                if (!(vdesc == null || vdesc == "" || vdesc == "0.00"))
                    vdesc = "\n" + "Valor Desconto: " + format(vdesc, 2);
                else
                    vdesc = "";
            }
            string txtAdic = vdesc + lacres;
            if (txtAdic != "")
            {
                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, txtAdic, ps);
                if (nroLinhasDescpro == 0)
                    return true; // dados adic não coube na pagina
            }

            return false; // hasMoreProdutos = false
        }

        protected virtual void SetValorISSQN(AcroFields af)
        {
            string v = GetValorXML("emit/IM");
            if (af.Fields.Contains(Getkey(af, "IM[0]")))
            {
                if (v == null)
                    v = "";                  
                af.SetField("IM", v);
            }

            v = GetValorXML("total/ISSQNtot/vServ");
            if (af.Fields.Contains(Getkey(af, "ISSQNtot[0]")))
            {
                if (v == null)
                    v = "";                  
                else
                    v = v.Replace('.', ',');
                af.SetField("ISSQNtot", v);
            }

            v = GetValorXML("total/ISSQNtot/vBC");
            if (af.Fields.Contains(Getkey(af, "vBC[0]")))
            {
                if (v == null)
                    v = "";                  
                else 
                    v = v.Replace('.', ',');
                af.SetField("vBC", v);
            }

            v = GetValorXML("total/ISSQNtot/vISS");
            if (af.Fields.Contains(Getkey(af, "vISSQN[0]")))
            {
                if (v == null)
                    v = "";                  
                else 
                    v = v.Replace('.', ',');
                af.SetField("vISSQN", v);
            }
        }

        protected virtual bool SetDadosAdicionais(PdfStamper ps, ref string infCpl)
        {
            // "xxxx xxxx xxxx /// xxxx xxxx xxxx /// xxxxx xxxxxx xxxxxxx /// xxxxxxxxxxxxxxxxxxx xxxxxxxxxxx /// xxxxxx"
            // "xxxx xxxx xxxx\nxxxx xxxx xxxx\nxxxxx xxxxxx xxxxxxx\nxxxxxxxxxxxxxxxxxxx xxxxxxxxxxx\nxxxxxx"

            if (infCpl != String.Empty)
            {
                infCpl = escreveDadosAdicionais(ps, "infCpl", infCpl);
            }

            if (infCpl == String.Empty)
                return false; // Não tem mais o que escrever.
            return true;
        }

        protected virtual string escreveDadosAdicionais(PdfStamper ps, string campo, string conteudo)
        {
            float[] pos = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, campo));
            PdfContentByte cb = ps.GetOverContent(1);  //writer.DirectContent;
            string infor = conteudo;
            while (true)
            {
                ColumnText ct = new ColumnText(cb);
                ct.SetSimpleColumn(pos[1] + 2, pos[2], pos[3] - 2, pos[4]);
                //ct.AdjustFirstLine = false;
                ct.Leading = font_infCpl.font.GetCalculatedLeading(1.1f);
                //ct.Alignment = Element.ALIGN_TOP;
                ct.Alignment = Element.ALIGN_BOTTOM;
                //ct.Alignment = Element.ALIGN_BASELINE;
                float Ypos = ct.YLine;

                Chunk p = new Chunk(infor, font_infCpl.font);
                //ColumnText ctx = ct;
                ct.AddText(p);
                if (!ColumnText.HasMoreText(ct.Go(true)))
                {
                    ct.YLine = Ypos;
                    ct.AddText(p);
                    ct.Go(false);
                    break;
                }
                int ultPos = infor.LastIndexOfAny(new char[] { ' ', '\n' });
                infor = infor.Remove(ultPos).Trim();
                //ct.ClearChunks();
                //ct.YLine = Ypos;
                //ct.FilledWidth = 0.0f;
            }
            //if (infor.Length < conteudo.Length)
            return conteudo.Remove(0, infor.Length).Trim();
        }

        protected virtual string GetDadosAdicionais()
        {
            StringBuilder infCpl = new StringBuilder("");
            //Complementares
            string v = GetValorXML("infAdic/infCpl");
            v = decode(v);

            //Contingência
            string tpEmis = GetValorXML("ide/tpEmis");
            if (tpEmis != "1")
            {
                Regex re = new Regex("conting.ncia", RegexOptions.IgnoreCase);
                if (! re.IsMatch(v))
                {
                    switch (tpEmis)
                    {
                        case "2":
                        case "5":
                            infCpl.Append("*** DANFE EM CONTINGÊNCIA, IMPRESSO EM DECORRÊNCIA DE PROBLEMAS TÉCNICOS. ***");
                            break;
                        case "4":
                            infCpl.Append("DANFE impresso em contingência - DPEC regularmente recebida pela Receita Federal do Brasil");
                            break;
                    }
                }
            }
            if ((!(arq.tpEmis == "1" || arq.tpEmis == "3")) && arq.status == "P")
            {
                PersistentCriteria pc = new PersistentCriteria("StatusNfei");
                pc.AddSelectEqualTo("nfei", nfei);
                pc.AddSelectEqualTo("StatusNfei[tipoStatusNfei].TipoStatusNfei[acao]", "AUTORIZACAO");
                pc.OrderBy("oid", TipoOrdenamento.Descendente);
                if (pc.Perform() > 0)
                {
                    StatusNfei sn = (StatusNfei)pc[0];
                    string protocolo = sn.numeroProtocolo.ToString();
                    infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                    infCpl.Append("*** NF-e Autorizada - Protocolo: " + protocolo + " - ");
                    infCpl.Append(sn.dataRecebimento.ToString("dd/MM/yyyy HH:mm:ss"));
                }
            }

            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(nfei.emitente, "Mensagem");
            string mensagem = "";
            if (pe != null)
                mensagem = pe.conteudo;
            if (!String.IsNullOrEmpty(mensagem))
            {
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
				infCpl.Append(mensagem);
            }
            //Retirada
            string retiradaCNPJ = GetValorXML("retirada/CNPJ");
            if ( ! String.IsNullOrEmpty(retiradaCNPJ))
            {
				infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append("Retirada: " + GetValorXML("retirada/xLgr") + ", " + GetValorXML("retirada/nro"));
                infCpl.Append(" " + GetValorXML("retirada/xCpl") + " - ");
                infCpl.Append(GetValorXML("retirada/xBairro") + " - " + GetValorXML("retirada/xMun") + "/" + GetValorXML("retirada/UF"));
            }
            //Entrega
            string entregaCNPJ = GetValorXML("entrega/CNPJ");
            if ( ! String.IsNullOrEmpty(entregaCNPJ))
            {
				infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append("Entrega: " + GetValorXML("entrega/xLgr") + ", " + GetValorXML("entrega/nro"));
                infCpl.Append(" " + GetValorXML("entrega/xCpl") + " - ");
                infCpl.Append(GetValorXML("entrega/xBairro") + " - " + GetValorXML("entrega/xMun") + "/" + GetValorXML("entrega/UF"));
            }

            //Complementares
            if (!String.IsNullOrEmpty(v))
            {
                v = v.Replace("||", "\n");
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append(v);
            }

            //Fisco
            v = GetValorXML("infAdic/infAdFisco");
			v = decode(v);
            if (!String.IsNullOrEmpty(v))
            {
                v = v.Replace("||", "\n");
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append(v);
            }
            return infCpl.ToString();
        }

        protected virtual int EscreverDescricao(int curlinha, int qtdlinhas, string descricao, PdfStamper ps)
        {
            float[] posFinal = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, "xProd[" + curlinha.ToString() + "]"));
            float[] posInicial = ps.AcroFields.GetFieldPositions(Getkey(ps.AcroFields, "xProd[" + (qtdlinhas - 1).ToString() + "]"));
            //float fieldPage = fieldPositions[0];
            //float fieldLlx = fieldPositions[1];
            //float fieldLly = fieldPositions[2];
            //float fieldUrx = fieldPositions[3];
            //float fieldUry = fieldPositions[4];
            PdfContentByte cb = ps.GetOverContent(1);  //writer.DirectContent;

            //Rectangle r = new Rectangle(posInicial[1], posInicial[2], posFinal[3], posFinal[4]);
            //r.BackgroundColor = Color.BLUE;
            //cb.Rectangle(r);
            //cb.Stroke();
            ColumnText ct = new ColumnText(cb);
            //ct.AdjustFirstLine = false;
            ct.SetSimpleColumn(posInicial[1], posInicial[2], posFinal[3], posFinal[4], font_xProd.font.GetCalculatedLeading(1.1f), Element.ALIGN_MIDDLE);
            //ct.Alignment = Element.ALIGN_TOP;
            //ct.Alignment = Element.ALIGN_BOTTOM;
            //ct.Alignment = Element.ALIGN_BASELINE;
            //ct.Alignment = Element.ALIGN_MIDDLE;
            
            Chunk p = new Chunk(descricao, font_xProd.font);
            if (font_xProd.textRise != 0f)
                p.SetTextRise(font_xProd.textRise);
            //Phrase p = new Phrase(descricao, font_xProd);
            //Paragraph p = new Paragraph(descricao, font_xProd);

            float Ypos = ct.YLine;
            ct.AddText(p);

            if (!ColumnText.HasMoreText(ct.Go(true)))
            {
                ct.YLine = Ypos;
                ct.AddText(p);
                ct.Go(false);
                return ct.LinesWritten;
            }
            return 0;

        }

        protected virtual myFont getFonte2(string defFonte)
        {
            //Formato : [Helv|Cour|Tmnr|Arial];size;textRise
            if (defFonte == String.Empty)
                defFonte = "helv;6.0;-0.5";
            string[] fonte_form = defFonte.Split(new char[] { ';' });
            float fontSize = float.Parse(fonte_form[1], ci_us);
            string fontName = fonte_form[0];
            BaseFont bf;
            if (fontName.ToUpper() == "COUR")
                bf = BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            else if (fontName.ToUpper() == "HELV")
                bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            else if (fontName.ToUpper() == "ARIAL")
                bf = BaseFont.CreateFont(Utils.GetPathAplicacao() + "Arial.ttf", BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            else // "TMNR"
                bf = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font f = new Font(bf, fontSize, Font.NORMAL);
            myFont mf = new myFont();
            mf.font = f;
            if (fonte_form.Length == 3)
                mf.textRise = float.Parse(fonte_form[2], ci_us);
            return mf;
        }
    }

    // DANFE V4
    public class DanfeV4 : DanfeV3
    {
        protected string versao;
        protected string versaoNFe;

        public DanfeV4(Nfei _nfei) : base(_nfei) 
        {
            versao = nfei.versao ?? "";
            versaoNFe = xmlNFE.Attributes["versao"].Value;
        }

        protected override void SetDestinatarioRemetente(AcroFields af)
        {
            SetMultValue(af, "dxNome", decode(GetValorXML("dest/xNome")));

            string v = GetValorXML("dest/CNPJ");
            if (v.Equals(""))
            {
                v = GetValorXML("dest/CPF");
                if (!v.Equals(""))
                    if (v.Length == 11)
                        v = Utils.Format(v, "cpf");
            }
            else if (v.Length == 14)
                v = Utils.FormatCNPJ(v);
            af.SetField("dCnpj", v);

            v = GetValorXML("ide/dEmi");
            string[] dts = v.Split(new char[] { '-' });
            v = dts[2] + "/" + dts[1] + "/" + dts[0];

            SetMultValue(af, "dEmi", v);

            v = GetValorXML("dest/enderDest/xLgr");
            v += ", " + GetValorXML("dest/enderDest/nro");
            string xCpl = GetValorXML("dest/enderDest/xCpl");
            if (xCpl.Length > 0)
                v += " /" + xCpl;
            af.SetField("dxLgr", v);

            v = GetValorXML("dest/enderDest/xBairro");
            af.SetField("dxBairro", v);

            v = GetValorXML("dest/enderDest/CEP");
            v = Utils.Format(v, "99.999-999");
            af.SetField("dCEP", v);

            v = GetValorXML("ide/dSaiEnt");

            if (v != null && v != "")
            {

                dts = v.Split(new char[] { '-' });
                v = dts[2] + "/" + dts[1] + "/" + dts[0];
                af.SetField("dSaiEnt", v);
            }

            //V4
            string horaSaida = GetValorXML("ide/hSaiEnt");
            if (String.IsNullOrEmpty(horaSaida))
            {

                XmlNode xmlObsCont = G6XmlUtils.FindXmlNodeByAttribute(xmlNFE, "infAdic/obsCont/xCampo", "HoraSaida");
                if (xmlObsCont != null)
                    horaSaida = xmlObsCont.InnerText ?? "";
            }
            af.SetField("horaSaida", horaSaida);

            v = GetValorXML("dest/enderDest/xMun");
            af.SetField("dxMun", v);

            v = GetValorXML("dest/enderDest/UF");
            af.SetField("dUF", v);

            v = GetValorXML("dest/IE");
            af.SetField("dIE", v);

            v = GetValorXML("dest/enderDest/fone");
            if (!v.Equals(""))
            {
                if (v.Length == 8)
                    v = Utils.Format(v, "9999-9999");
                else if (v.Length == 10)
                    v = Utils.Format(v, "(99)9999-9999");
            }

            af.SetField("dfone", v);
        }

        private string formatVar(string valor, bool milhar)
        {
            string strDec = milhar ? "#,##0" : "###0";
            Decimal d = Convert.ToDecimal(valor, ci_us);
            string[] a = valor.Split('.');
            if (a.Length > 1)
            {
                strDec += "." + Regex.Replace(a[1], @"\d", "0");
            }
            return d.ToString(strDec, ci_br);
        }

        protected override bool SetDadosProduto(PdfStamper ps, ref int proximoProduto)
        {
            AcroFields af = ps.AcroFields;
            int qtdLinhas = 0;
            int nroLinhasDescpro;

            for (int i = 0; i < 200; i++)
            {
                string test = "cProd[" + i.ToString() + "]";
                if (af.Fields.Contains(Getkey(af, test)))
                    continue;
                qtdLinhas = i;
                break;
            }

            int curLinha = 0;

            ArrayList ar = G6XmlUtils.GetAllXmlNodes(xmlNFE, "prod");
            for (int i = proximoProduto; i < ar.Count; i++, proximoProduto++)
            {
                if (curLinha >= qtdLinhas)
                    return true;  // Acabou espaço nesta página, pegue outra...

                XmlNode xn = (XmlNode)ar[i];
                string codProd = G6XmlUtils.GetInnerXml(xn, "cProd");
                string descprod = G6XmlUtils.GetInnerXml(xn, "xProd");
                XmlNode xinf = G6XmlUtils.FindXmlNode(xn.ParentNode, "infAdProd");
                if (xinf != null)
                    descprod = descprod.Trim() + "||" + xinf.InnerText;
                descprod = descprod.Replace("||", "\n");
                descprod = descprod.Trim();
                descprod = decode(descprod);

                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, descprod, ps);
                if (nroLinhasDescpro == 0)
                    return true; // tem mais produtos (não coube na pagina)

                string ncm = G6XmlUtils.GetInnerXml(xn, "NCM");
                // V4 - mantido pois pode imprimir DANFES da v1
                if (String.IsNullOrEmpty(ncm))
                    ncm = G6XmlUtils.GetInnerXml(xn, "genero");
                string cfop = G6XmlUtils.GetInnerXml(xn, "CFOP");
                string uni = G6XmlUtils.GetInnerXml(xn, "uCom");
                string qtde = G6XmlUtils.GetInnerXml(xn, "qCom") ?? "0";
                qtde = formatVar(qtde, false); // V4

                string vlunit = G6XmlUtils.GetInnerXml(xn, "vUnCom") ?? "0";
                vlunit = formatVar(vlunit, true); // V4

                string vltotal = G6XmlUtils.GetInnerXml(xn, "vProd") ?? "0";
                vltotal = format(vltotal, 2);

                ArrayList ai = G6XmlUtils.GetAllXmlNodes(xmlNFE, "imposto");
                XmlNode xi = (XmlNode)ai[i];

                string orig = null;
                string cst = "";
				string csosn = "";
                string bc = "";
                string icms = "";
                string bcst = "";
                string icmsst = "";
                string ipi = "";
                string aliqicms = "";
                string aliqipi = "";

				XmlNode nICMS =  G6XmlUtils.FindXmlNode(xi, "ICMS");
				if (nICMS != null)
				{
					XmlNode tICMS = nICMS.FirstChild;
					orig = G6XmlUtils.GetInnerXml(tICMS, "orig") ?? "";
					cst = G6XmlUtils.GetInnerXml(tICMS, "CST") ?? "";
					bc = G6XmlUtils.GetInnerXml(tICMS, "vBC") ?? "";
					icms = G6XmlUtils.GetInnerXml(tICMS, "vICMS") ?? "";
					aliqicms = G6XmlUtils.GetInnerXml(tICMS, "pICMS") ?? "";
					bcst = G6XmlUtils.GetInnerXml(tICMS, "vBCST") ?? "";
					icmsst = G6XmlUtils.GetInnerXml(tICMS, "vICMSST") ?? "";

					csosn = G6XmlUtils.GetInnerXml(tICMS, "CSOSN") ?? "";
				}


                if (af.Fields.Contains(Getkey(af, "vIPI[0]")))
                {
                    ipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/vIPI") ?? "";
	                ipi = format(ipi, 2);
                }
                if (af.Fields.Contains(Getkey(af, "pIPI[0]")))
                {
                    aliqipi = G6XmlUtils.GetInnerXml(xi, "IPI/IPITrib/pIPI") ?? "";
                    aliqipi = format(aliqipi, 2);
                }

                bc = format(bc, 2);
                aliqicms = format(aliqicms, 2);
                icms = format(icms, 2);
                bcst = format(bcst, 2);
                icmsst = format(icmsst, 2);

                string vlinha = "[" + curLinha.ToString() + "]";

                af.SetField("cProd" + vlinha, codProd);
                af.SetField("NCM" + vlinha, ncm);
				// Novo na V2.00 -  CSOSN
				if (af.Fields.Contains(Getkey(af, "CST[0]")))
					af.SetField("CST" + vlinha, orig + cst);
				if (af.Fields.Contains(Getkey(af, "CSOSN[0]")))
					af.SetField("CSOSN" + vlinha, orig + csosn);
				af.SetField("CFOP" + vlinha, cfop);
                af.SetField("uCom" + vlinha, uni);
                af.SetField("qCom" + vlinha, qtde);
                af.SetField("vUnCom" + vlinha, vlunit);
                af.SetField("vProd" + vlinha, vltotal);
                af.SetField("vBCICMS" + vlinha, bc);
                af.SetField("vICMS" + vlinha, icms);
                //SUBST. TRIB
                if (af.Fields.Contains(Getkey(af, "vBCST[0]")))
                    af.SetField("vBCST" + vlinha, bcst);
                if (af.Fields.Contains(Getkey(af, "vICMSST[0]")))
                    af.SetField("vICMSST" + vlinha, icmsst);
                //IPI Opcional
                if (af.Fields.Contains(Getkey(af, "vIPI[0]")))
                    af.SetField("vIPI" + vlinha, ipi);
                af.SetField("pICMS" + vlinha, aliqicms);
                if (af.Fields.Contains(Getkey(af, "pIPI[0]")))
                    af.SetField("pIPI" + vlinha, aliqipi);

                curLinha = curLinha + nroLinhasDescpro;
            }

            ArrayList arL = G6XmlUtils.GetAllXmlNodes(xmlNFE, "lacres");
            string lacres = "";
            for (int i = 0; i < arL.Count; i++)
            {
                XmlNode xn = (XmlNode)arL[i];
                string lacre = G6XmlUtils.GetInnerXml(xn, "nLacre");
                lacres = lacres + lacre + ", ";
            }
            //string lacres = GetValorXML("transp/vol/lacres/nLacre");
            if (!(lacres == null || lacres == ""))
                lacres = "\n" + "Lacres: " + lacres;

            string vdesc = "";
            if (!hasFieldDesconto)
            {
                vdesc = GetValorXML("total/ICMSTot/vDesc");
                if (!(vdesc == null || vdesc == "" || vdesc == "0.00"))
                    vdesc = "\n" + "Valor Desconto: " + format(vdesc, 2);
                else
                    vdesc = "";
            }
            string txtAdic = vdesc + lacres;
            if (txtAdic != "")
            {
                nroLinhasDescpro = EscreverDescricao(curLinha, qtdLinhas, txtAdic, ps);
                if (nroLinhasDescpro == 0)
                    return true; // dados adic não coube na pagina
            }

            return false; // hasMoreProdutos = false
        }

        protected override string GetDadosAdicionais()
        {
            StringBuilder infCpl = new StringBuilder("");
            string v = GetValorXML("infAdic/infCpl");
            v = decode(v);

            string tpEmis = GetValorXML("ide/tpEmis");
            if (tpEmis != "1")
            {
                //Se não contiver o texto de "Contingência", adiciona
                Regex re = new Regex("conting.ncia", RegexOptions.IgnoreCase);
                if (!re.IsMatch(v))
                {
                    switch (tpEmis)
                    {
                        case "2":
                        case "5":
                            infCpl.Append("*** DANFE EM CONTINGÊNCIA, IMPRESSO EM DECORRÊNCIA DE PROBLEMAS TÉCNICOS. ***");
                            break;
                        case "4":
                            infCpl.Append("DANFE impresso em contingência - DPEC regularmente recebida pela Receita Federal do Brasil");
                            break;
                    }
                }
                string dhCcont = GetValorXML("ide/dhCont");
                if (!String.IsNullOrEmpty(dhCcont))
                {
                    infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                    infCpl.Append("*** Contingência - Motivo: " + GetValorXML("ide/xJust"));
                    DateTime dt = DateTime.Parse(dhCcont, null, System.Globalization.DateTimeStyles.AssumeLocal);
                    infCpl.Append(", Data/hora entrada: " + dt.ToString("dd/MM/yyyy HH:mm:ss") + " ***");
                }
            }

            if ((!(arq.tpEmis == "1" || arq.tpEmis == "3")) && arq.status == "P")
            {
                PersistentCriteria pc = new PersistentCriteria("StatusNfei");
                pc.AddSelectEqualTo("nfei", nfei);
                pc.AddSelectEqualTo("StatusNfei[tipoStatusNfei].TipoStatusNfei[acao]", "AUTORIZACAO");
                pc.OrderBy("oid", TipoOrdenamento.Descendente);
                if (pc.Perform() > 0)
                {
                    StatusNfei sn = (StatusNfei)pc[0];
                    string protocolo = sn.numeroProtocolo.ToString();
                    infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                    infCpl.Append("*** NF-e Autorizada - Protocolo: " + protocolo);
                    infCpl.Append(" - " + sn.dataRecebimento.ToString("dd/MM/yyyy HH:mm:ss") + " ***");
                }
            }

            ParametroEmitente pe = ParametroEmitente.GetParametroEmitentePorEmitenteENome(nfei.emitente, "Mensagem");
            string mensagem = "";
            if (pe != null)
                mensagem = pe.conteudo;
            if (!String.IsNullOrEmpty(mensagem))
            {
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append(mensagem);
            }

			// Novo V2 - CPF ou CNPJ
			//Retirada
            //string retiradaCNPJ = GetValorXML("retirada/CNPJ");
			//if (!String.IsNullOrEmpty(retiradaCNPJ))
			XmlNode retEnt = G6XmlUtils.FindXmlNode(xmlNFE, "retirada");
			if (retEnt != null)
            {
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
				infCpl.Append(String.Format("Retirada: Cpf/Cnpj {0} Local: ", Utils.FormatCPFCNPJ(retEnt.FirstChild.InnerText)));
                infCpl.Append( GetValorXML("retirada/xLgr") + ", " + GetValorXML("retirada/nro"));
                infCpl.Append(" " + GetValorXML("retirada/xCpl") + " - ");
                infCpl.Append(GetValorXML("retirada/xBairro") + " - " + GetValorXML("retirada/xMun") + "/" + GetValorXML("retirada/UF"));
            }
			// Novo V2 - CPF ou CNPJ
			//Entrega
			//string entregaCNPJ = GetValorXML("entrega/CNPJ");
			//if (!String.IsNullOrEmpty(entregaCNPJ))
			retEnt = G6XmlUtils.FindXmlNode(xmlNFE, "entrega");
			if (retEnt != null)
			{
				infCpl.Append(infCpl.Length == 0 ? "" : "\n");
				infCpl.Append(String.Format("Entrega: Cpf/Cnpj {0} Local: ", Utils.FormatCPFCNPJ(retEnt.FirstChild.InnerText))); 
				infCpl.Append(GetValorXML("entrega/xLgr") + ", " + GetValorXML("entrega/nro"));
				infCpl.Append(" " + GetValorXML("entrega/xCpl") + " - ");
				infCpl.Append(GetValorXML("entrega/xBairro") + " - " + GetValorXML("entrega/xMun") + "/" + GetValorXML("entrega/UF"));
			}

            //Complementares
            if (!String.IsNullOrEmpty(v))
            {
                v = v.Replace("||", "\n");
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append(v);
            }

            //Fisco
            v = GetValorXML("infAdic/infAdFisco");
            v = decode(v);
            if (!String.IsNullOrEmpty(v))
            {
                v = v.Replace("||", "\n");
                infCpl.Append(infCpl.Length == 0 ? "" : "\n");
                infCpl.Append(v);
            }
            return infCpl.ToString();
        }

        protected override void SetTransportadora(AcroFields af)
        {
            string v = decode(GetValorXML("transp/transporta/xNome"));
            //af.SetField("txNome", v);
            SetMultValue(af, "txNome", v);

            //V4
            v = GetValorXML("transp/modFrete");
            //af.SetField("modFrete", v);

            string val = "";
            switch (v)
            {
                case "0":
                    val = "Emitente";
                    break;
                case "1":
                    val = "Dest/Rem";
                    break;
                case "2":
                    val = "Terceiros";
                    break;
                case "9":
                    val = "Sem Frete";
                    break;
            }
            af.SetField("xModFrete", val);
			af.SetField("modFrete", v);

            v = GetValorXML("transp/veicTransp/placa");
            af.SetField("tPlacaVeiculo", v);

            v = GetValorXML("transp/veicTransp/UF");
            af.SetField("tvUF", v);

            v = GetValorXML("transp/veicTransp/RNTC");
            af.SetField("rntc", v);

            v = GetValorXML("transp/transporta/CNPJ");
            v = Utils.FormatCNPJ(v);
            af.SetField("tCnpj", v);

            v = decode(GetValorXML("transp/transporta/xEnder"));
            af.SetField("txLgr", v);

            v = GetValorXML("transp/transporta/xMun");
            af.SetField("txMun", v);

            v = GetValorXML("transp/transporta/UF");
            af.SetField("tUF", v);

            v = GetValorXML("transp/transporta/IE");
            af.SetField("tIE", v);

            v = GetValorXML("transp/vol/qVol");
            af.SetField("qVol", v);

            v = decode(GetValorXML("transp/vol/esp"));
            af.SetField("esp", v);

            v = decode(GetValorXML("transp/vol/marca"));
            af.SetField("marca", v);

            v = GetValorXML("transp/vol/nVol");
            af.SetField("nVol", v);

            v = GetValorXML("transp/vol/pesoB");
            v = v.Replace('.', ',');
            af.SetField("pesoB", v);

            v = GetValorXML("transp/vol/pesoL");
            v = v.Replace('.', ',');
            af.SetField("pesoL", v);
        }

    }

}
