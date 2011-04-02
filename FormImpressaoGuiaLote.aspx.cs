using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Threading;
using ObjetosDeNegocio;
using System.IO;
using System.Web.SessionState;

namespace DesWebServer
{
    public partial class FormImpressaoGuiaLote : G6Form
    {
        protected new void Page_Load(object sender, EventArgs e)
        {
            CarregarG6FormHeader((G6FormHeader)pnlHeader.FindControl("G6Header"));
            EsconderBotaoPesquisar();
            DisableTimeout();
            if (g6Progress.isActive)
                g6Progress.OnProcessEnd = this.OnProcessEnd;
            if (!IsPostBack)
            {
                PreencherProcessamentos();
                CarregarConvenio();
                PreencherRegimeISS();
                PreencherConta();
                if (Session["ThreadGerarGuias"] != null && ((Thread)Session["ThreadGerarGuias"]).IsAlive)
                {
                    btnGerarGuias.Enabled = false;
                    g6Progress.AbrirPainelProgresso("Gerar Guias em Lote", "ProgressoGuiasLote", true);
                    MessageBox("Aguarde, há um processo em andamento.");
                }
                else
                {
                    btnGerarGuias.Enabled = true;
                    pnlProgresso.Visible = false;
                }
                btnMostrarGrid.Enabled = false;
            }

            if (GetDireitosDoUsuarioLogado("EditarQueryNaImpressaoDeGuias") != "")
            {
                laQuery.Visible = true;
                edQuery.Visible = true;
            }
                            
            base.Page_Load(sender, e);

        }
        private void PreencherRegimeISS()
        {
            PersistentCriteria pc = new PersistentCriteria("TipoRegimeIss", ddlRegimeISS, "descricao");
            ddlRegimeISS.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Selecione Caso Necessário --"));
        }

        private void PreencherProcessamentos()
        {
            PersistentCriteria pc = new PersistentCriteria("Processamento");
            pc.AddSelectIn("tipoProcessamento", new string[] { "PDGuiaDiaria", "PMGuiaMensal", "PAGuiaAnual" });
            pc.OrderBy("oid", TipoOrdenamento.Descendente);
            pc.SetMaxRegistros(10);
            ddlProcessamentos.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Ultimos Processamentos de Guias --", "-1"));
            if (pc.Perform() > 0)
            {
                foreach(Processamento proc in pc)
                {
                    ddlProcessamentos.Items.Add(new ListItem(proc.oid.ToString() + " - " + proc.dataProcessamento.ToString("dd/MM/yyyy")+" - "+proc.tipoProcessamento, proc.oid.ToString()));
                }
            }
        }

        private void PreencherConta()
        {
            PersistentCriteria pc = new PersistentCriteria("TipoConta", ddlConta, "descricao");
            ddlConta.Items.Insert(0, new System.Web.UI.WebControls.ListItem("-- Selecione Caso Necessário --"));
        }

        private void CarregarConvenio()
        {
            PersistentCriteria pc = new PersistentCriteria("ConvenioBanco");
            if (pc.Perform() > 0)
            {
                foreach (ConvenioBanco cb in pc)
                {
                    cb.Retrieve();
                    string texto = cb.banco.codBanco + "-" + cb.banco.nome + "-" + cb.codCarteira;
                    ddlBanco.Items.Add(new System.Web.UI.WebControls.ListItem(texto, cb.oid.ToString()));
                }
                ddlBanco.SelectedIndex = 0;
            }
        }

        private bool VerificarParametrosImpressao()
        {
            if (((edGuiaInicial.Text.Trim() != "") && (edGuiaFinal.Text.Trim() == "")) ||
                ((edGuiaInicial.Text.Trim() == "") && (edGuiaFinal.Text.Trim() != "")))
            {
                MessageBox("Favor informar o número de guia início e o número da guia final corretamente!");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Função de Callback de retorno para apresentacao do contador
        /// </summary>
        /// <param name="_controleProgresso"></param>
        public void ApresentaDadosParaUsuario(ControleProgresso _controleProgresso)
        {
            Session["ProgressoGuiasLote"] = _controleProgresso;
        }

        public bool UsarQuery(string query)
        {
            if (query.ToUpper().IndexOf("INSERT") > 0 || query.ToUpper().IndexOf("UPDATE") > 0 || query.ToUpper().IndexOf("DELETE") > 0)
                return false;

            if (query.ToUpper().IndexOf("...") > 0 || query.ToUpper().IndexOf("WHERE") == 0)
                return false;
            else
            return true;
        }

        public DataTable GetGuiasPorQuery(string query)
        {
            DataTable dt;
            try
            {
                dt = PersistentBroker.singleInstance.ProcessSql(query);
                
                if (dt != null)
                    return dt;
                else return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        protected void btnGerarGuias_Click(object sender, System.EventArgs e)
        {
            ArrayList guias;

            if (UsarQuery(edQuery.Text))
            {
                DataTable dt = GetGuiasPorQuery(edQuery.Text);

                if (dt != null)
                {
                    guias = new ArrayList();

                    foreach(DataRow dr in dt.Rows)
                        guias.Add(new GuiaRecolhimento(Convert.ToInt32(dr["oid"])));
                }
                else
                {
                    MessageBox("A query utilizada não compila corretamente. Verifique.");
                    return;
                }
            }
            else
            {
                if (!VerificarParametrosImpressao())
                    return;

                if (Session["ThreadGerarGuias"] != null)
                {
                    if (((Thread)Session["ThreadGerarGuias"]).IsAlive)
                    {
                        MessageBox("Aguarde, o processamento está em andamento.");
                        return;
                    }
                }


                PersistentCriteria pc = new PersistentCriteria("GuiaRecolhimento");
                if (chkSomenteNaoImpressas.Checked)
                {
                    pc.AddSelectEqualTo("status", "G");
                }
                else
                {
                    pc.AddSelectIn("status", new string[] { "G", "N" });
                }

                if (edDtVencimentoInicial.Text.Trim() != "" && edDtVencimentoFinal.Text.Trim() != "")
                {
                    pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "dataVencimento", Utils.StringToDate(edDtVencimentoInicial.Text)));
                    pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "dataVencimento", Utils.StringToDate(edDtVencimentoFinal.Text)));
                }
                else
                {
                    if (edDtVencimentoInicial.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "dataVencimento", Utils.StringToDate(edDtVencimentoInicial.Text)));
                    }
                    if (edDtVencimentoFinal.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "dataVencimento", Utils.StringToDate(edDtVencimentoFinal.Text)));
                    }
                }

                //dataEmissao
                if (edDtEmissaoInicial.Text.Trim() != "" && edDtEmissaoFinal.Text.Trim() != "")
                {
                    pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "dataEmissao", Utils.StringToDate(edDtEmissaoInicial.Text)));
                    pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "dataEmissao", Utils.StringToDate(edDtEmissaoFinal.Text)));
                }
                else
                {
                    if (edDtEmissaoInicial.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "dataEmissao", Utils.StringToDate(edDtEmissaoInicial.Text)));
                    }
                    if (edDtEmissaoFinal.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "dataEmissao", Utils.StringToDate(edDtEmissaoFinal.Text)));
                    }
                }

                if (((edGuiaInicial.Text.Trim() != "") && (edGuiaFinal.Text.Trim() != "")))
                {
                    pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "numGuia", Convert.ToInt32(edGuiaInicial.Text)));
                    pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "numGuia", Convert.ToInt32(edGuiaFinal.Text)));
                }
                else
                {
                    if (edGuiaInicial.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.GreaterOrEqualThan, "numGuia", Convert.ToInt32(edGuiaInicial.Text)));
                    }
                    if (edGuiaFinal.Text.Trim() != "")
                    {
                        pc.AddCriteria(new Criteria(TipoCriteria.LessOrEqualThan, "numGuia", Convert.ToInt32(edGuiaFinal.Text)));
                    }
                }


                if (edIdProcessamento.Text.Trim() != "")
                {
                    try
                    {
                        int idproc = Convert.ToInt32(edIdProcessamento.Text.Trim());
                        pc.AddSelectEqualTo("processamento", new Processamento(idproc));
                    }
                    catch
                    {
                        MessageBox("Número do processamento informado incorretamente!");
                        return;
                    }
                }

                if (edInscricao.Text.Trim() != "")
                {
                    Contribuinte c = Contribuinte.GetContribuinteByInscricao(edInscricao.Text.Trim());
                    if (c == null)
                    {
                        MessageBox("Inscrição do contribuinte informado não encontrado!");
                        return;
                    }
                    pc.AddSelectEqualTo("contribuinte", c);
                }

                if (ddlRegimeISS.SelectedIndex > 0)
                    pc.AddSelectEqualTo("GuiaRecolhimento[contribuinte].Contribuinte[tipoRegimeIss]", new TipoRegimeIss(Convert.ToInt32(ddlRegimeISS.SelectedValue)));

                if (ddlConta.SelectedIndex > 0)
                {
                    pc.AddSelectEqualTo("tipoConta", new TipoConta(Convert.ToInt32(ddlConta.SelectedValue)));
                }

                if (edValorMinimo.Text.Trim() != "")
                {
                    pc.AddSelectGreaterThan("valorGuia", Decimal.Parse(edValorMinimo.Text));
                }

                pc.OrderBy("numGuia", TipoOrdenamento.Ascendente);
                pc.Perform();
                guias = pc.GetAllObjects();
            }

            if (guias.Count > 0)
            {
                btnGerarGuias.Enabled = false;
                edQuery.Text = "Recarregue o programa para resetar a query.";

                ConvenioBanco cb = new ConvenioBanco(Convert.ToInt32(ddlBanco.SelectedValue));
                cb.Retrieve();


                ThreadGerarGuias tc = new ThreadGerarGuias(this, GetUsuarioCorrente(), guias, cb, Utils.GetPathAplicacao(), Session);
                Thread tExecutarCalculo = new Thread(new ThreadStart(tc.GerarGuias));
                tExecutarCalculo.Start();
                Session["ThreadGerarGuias"] = tExecutarCalculo;
                g6Progress.AbrirPainelProgresso("Gerar Guias em Lote", "ProgressoGuiasLote", true);
                g6Progress.OnProcessEnd = this.OnProcessEnd;
            }
            else
                MessageBox("O critério de seleção não encontrou nenhuma guia de recolhimento para ser impressa!");

        }

        private void ItemCommand(object source, System.Web.UI.WebControls.DataGridCommandEventArgs e)
        {
            if (e.CommandName.Equals("Imprimir"))
            {
                Thread.Sleep(2000);
                string dir = Parametro.GetParametroPorNome("NomeDiretorioGeraGuias");
                if (dir.Equals(""))
                    dir = "Temp";
                dir = Utils.GetRelativePath(Page.Request) + "/" + dir + "/" + e.Item.Cells[2].Text;
                AbrirJanelaModalDeRelatorioHtml(dir, null);
            }
        }

        protected void btnMostrarGrid_Click(object sender, System.EventArgs e)
        {
            if (Session["gridGuias"] != null)
            {
                grdGuias.Visible = true;
                grdGuias.DataSource = (DataTable)Session["gridGuias"];
                grdGuias.DataBind();
            }
        }

        protected void btnAtualizar_Click(object sender, EventArgs e)
        {
        }

        protected void grdGuias_ItemCommand(object source, DataGridCommandEventArgs e)
        {
            if (e.CommandName.Equals("Imprimir"))
            {
                Thread.Sleep(2000);
                string dir = Parametro.GetParametroPorNome("NomeDiretorioGeraGuias");
                if (dir.Equals(""))
                    dir = "Temp";
                dir = Utils.GetRelativePath(Page.Request) + "/" + dir + "/" + e.Item.Cells[2].Text;
                AbrirJanelaModalDeRelatorioHtml(dir, null);
            }
        }

        public void ProcessTerminated(string message, bool success)
        {
            btnMostrarGrid.Enabled = true;
            //btnMostrarGrid_Click(null, null);
            g6Progress.ProcessTerminated(message, success);

        }

        public void OnProcessEnd(bool Success)
        {
            btnMostrarGrid.Enabled = true;
            //btnMostrarGrid_Click(null, null);
        }

        protected override void RaisePostBackEvent(IPostBackEventHandler sourceControl, string eventArgument)
        {
            g6Progress.OnProcessEnd = this.OnProcessEnd;
            base.RaisePostBackEvent(sourceControl, eventArgument);
        }

        protected void ddlProcessamentos_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt32(ddlProcessamentos.SelectedValue) > 0)
            {
                edIdProcessamento.Text = ddlProcessamentos.SelectedValue;
            }
        }

        public void HabilitarBotaoGridGuias()
        {
            btnMostrarGrid.Enabled = true;
        }
    }

    #region ThreadGerarGuias
    public class ThreadGerarGuias
    {
        ArrayList pc;
        FormImpressaoGuiaLote formImpressaoGuiaLote;
        ConvenioBanco cb;
        C1.C1PrintDocument.C1PrintDocument repGuia;
        string pathAplicacao;
        Usuario usuario;
        HttpSessionState Session;


        public ThreadGerarGuias(FormImpressaoGuiaLote _formImpressaoGuiaLote, Usuario _usuario, ArrayList _pc, ConvenioBanco _cb, string _pathAplicacao, HttpSessionState _session)
        {
            usuario = _usuario;
            formImpressaoGuiaLote = _formImpressaoGuiaLote;
            pc = _pc;
            cb = _cb;
            repGuia = new C1.C1PrintDocument.C1PrintDocument();
            repGuia.PageSettings.PrinterSettings.PrinterName = "";
            repGuia.C1DPageSettings = "color:False;landscape:False;margins:100,100,100,100;papersize:827,1169,QQA0AA==";
            repGuia.ColumnSpacingStr = "0.5in";
            repGuia.ColumnSpacingUnit.DefaultType = true;
            repGuia.ColumnSpacingUnit.UnitValue = "0.5in";
            repGuia.DefaultUnit = C1.C1PrintDocument.UnitTypeEnum.Inch;
            repGuia.DocumentName = "";
            pathAplicacao = _pathAplicacao;
            Session = _session;
        }

        public void GerarGuias()
        {
            GerarGuiaLote pa = new GerarGuiaLote();
            G6Session gs = new G6Session(usuario, true);
            try
            {
                pa.SetCallBack(new CallBackProcessamento(formImpressaoGuiaLote.ApresentaDadosParaUsuario));
                pa.SetControleProgresso(criarControleProgresso());
                pa.pc = pc;
                pa.usuario = usuario;
                pa.formImpressaoGuiaLote = formImpressaoGuiaLote;
                pa.repGuia = repGuia;
                pa.pathAplicacao = pathAplicacao;
                pa.cb = cb;
                pa.CalcularTodos(gs);
                Session["gridGuias"] = pa.GetGridGuias();
                formImpressaoGuiaLote.HabilitarBotaoGridGuias();
                formImpressaoGuiaLote.ProcessTerminated("Fim da Geração de Guias em Lote", true);
                
                //COMMIT
                gs.CommitTransaction();
            }
            catch (Exception ex)
            {
                //ROLLBACK
                if (gs.InTransaction())
                    gs.RollbackTransaction();                
                formImpressaoGuiaLote.ProcessTerminated("Erro na Geração de Guias em Lote: " + ex.Message + " STACK TRACE: " + ex.StackTrace, false);
                Utils.RegisterLogEvento(ex, "Erro ao gerar guias em lote", "GuiaRecolhimento", Utils.EventType.Error, usuario);
            }
        }
        private ControleProgresso criarControleProgresso()
        {
            ControleProgresso cp = new ControleProgresso();
            cp.descricao = "Geração de Guia em Lote";
            cp.id = "GERARGUIALOTE";
            cp.posicaoAtual = 0;
            cp.qtdeTotal = 0;
            cp.tempoDecorrido = "00:00:00";
            cp.tempoInicio = DateTime.Now;
            cp.tempoTermino = DateTime.Now;
            cp.status = "N";
            cp.Save();
            return cp;
        }
    }
    #endregion ThreadGerarGuias

}
