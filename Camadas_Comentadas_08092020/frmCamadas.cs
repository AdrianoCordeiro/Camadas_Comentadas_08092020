using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Camada_Negocio_Preferencia_BLL;
using Preferencia_Model_VO;
using Excel = Microsoft.Office.Interop.Excel; // criacao de alias para o ms excel
using Email = Microsoft.Office.Interop.Outlook; // criação de alias para o ms outlook

namespace Camadas_Comentadas_08092020
{
    public partial class frmCamadas : Form
    {
        Preferencia objPreferencia;
        Familiares objFamiliar;
        PrefFamBLL objPrefFamBLL;
        PreferenciaVO objPreferenciaVO;
        FamiliaresVO objFamiliarVO;
        PrefFamVO objPrefFamVO;
        int intValorAntigo,intCodFamiliarAntigo, intPrefFamId, IntPrefFamCod;
        string strValorAntigo, strValorAntigoPrefFamNome, strValorAntigoDescricao;
        bool boolValorInserido, boolValorInseridoFam, boolInserirPrefFam;
       
        Excel._Application objExcelApplicacao;// objetos de excel
        Excel.Workbook objExcelArquivo;// objetos de excel sem underline
        Excel.Worksheet objExcelPlanilha;// objetos de excel sem underline
        Excel.Range objExcelPlanilhaCabecalho;
        Excel.Range objExcelPlanilhaDados;
        
        Email.Application objEmailApp;
        Email.MailItem objMensagem;
        Email.OlAttachmentType objArquivoAnexoTipo;

        string[] arrayEmailArqAnexo = new String[0];
        long lngEmailArqAnexoPosicao;
        string strDisplayName;

        public frmCamadas()
        {
            InitializeComponent();
        }

// -------------------------------------PREFERENCIAS-------------------------------------------------------
        private void btnIfElse_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Escolha Entre Ok ou Cancelar!", "Aviso!", MessageBoxButtons.OKCancel)== System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show("Voce Escolheu Ok!"); 
            }
            else
            {
                MessageBox.Show("Voce Escolheu Cancelar!"); 
            }
        }

        private void btnImpTxtWhile_Click(object sender, EventArgs e)
        {
            objPreferencia = new Preferencia();

            lstbxPreferencias.Items.Clear();

            lstbxPreferencias.Items.AddRange(objPreferencia.ImportaTextoWhile().ToArray());
        }

        private void btnImpBd_Click(object sender, EventArgs e)
        {
            objPreferencia = new Preferencia();

            lstbxPreferencias.Items.Clear();

            lstbxPreferencias.Items.AddRange(objPreferencia.ImportarBdConectado().ToArray());
        }

        private void btnImpBdDesc_Click(object sender, EventArgs e)
        {
            objPreferencia = new Preferencia();

            lstbxPreferencias.Items.Clear();

            lstbxPreferencias.Items.AddRange(objPreferencia.ImportarBdDesconectado().ToArray());
        }

        private void btnConsBd_Click(object sender, EventArgs e)
        {
            ConsultarBd();
        }
        public void ConsultarBd(int ? intId  = null, string strPreferencia = null)
        {
            try // a funcao try acopla os comandos da funcao
            {
                objPreferenciaVO = new PreferenciaVO(); // sem parametros

                if (!string.IsNullOrEmpty(intId.ToString()))
                {
                    objPreferenciaVO.setId(Convert.ToInt32(intId));
                }
                // trabalham independentemente primeiro perguntando pelo id e depois pela descricao
                if (!string.IsNullOrEmpty(strPreferencia))
                {
                    objPreferenciaVO.setDescricao(strPreferencia);//seter classico ao inves do setter da ms
                    //   objPreferenciaVO.Descricao = strPreferencias; // Setter da Ms do Setter classico
                }

                //construcao da model com a sobrecarga com parametros 
                //objPreferenciaVO = new PreferenciasVO(strPreferencias);//com parametro nunca chamar se null

                objPreferencia = new Preferencia();

                bndsrcPreferencias.DataSource = objPreferencia.ConsultarBd(objPreferenciaVO);

                dtgwPreferencias.DataSource = bndsrcPreferencias;

            }
            catch (Exception ex)
            {

                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        private void frmCamadas_Load(object sender, EventArgs e)
        {
            ConsultarBd();
            ConsultarBdFam();
            dtgvvwPrefamRefresh();
        }

        private void dtgwPreferencias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strValorAntigo = dtgwPreferencias.CurrentRow.Cells["Descricao"].Value.ToString();
            if (!string.IsNullOrEmpty(dtgwPreferencias.CurrentRow.Cells["ID"].Value.ToString()))
            {
                intValorAntigo = Convert.ToInt32(dtgwPreferencias.CurrentRow.Cells["ID"].Value.ToString());
            }
        }

        private void btnIncBd_Click(object sender, EventArgs e)
        {
            IncluirBd(dtgwPreferencias.CurrentCell.EditedFormattedValue.ToString());
        }
        public void IncluirBd(string strPreferencia)
        {
            try
            {
                objPreferenciaVO = new PreferenciaVO(strPreferencia);
              
                objPreferencia = new Preferencia();

                if (objPreferencia.IncluirBd(objPreferenciaVO))
                {
                    MessageBox.Show("Inclusao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Inclusao!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        private void btnExcBd_Click(object sender, EventArgs e)
        {
            ExcluirBd(Convert.ToInt32(dtgwPreferencias.CurrentRow.Cells["ID"].Value.ToString()));
        }
        public void ExcluirBd(int intId)
        {
            try
            {
                objPreferenciaVO = new PreferenciaVO();
                objPreferenciaVO.setId(intId);
                objPreferencia = new Preferencia();

                if (objPreferencia.ExcluirBd(objPreferenciaVO))
                {
                    MessageBox.Show("Exclusao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Exclusao!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        private void btnAltBd_Click(object sender, EventArgs e)
        {
            AlterarBd(intValorAntigo, dtgwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
        }
        public void AlterarBd(int intIdPreferencia, string strNovo)
        {
            try
            {
                objPreferenciaVO = new PreferenciaVO(intIdPreferencia, strNovo);
                objPreferencia = new Preferencia();

                if (objPreferencia.AlterarBd(objPreferenciaVO))
                {
                    MessageBox.Show("Alteracao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Alteracao!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            boolValorInserido = true;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir " + strValorAntigo, "Aviso", MessageBoxButtons.YesNo)== System.Windows.Forms.DialogResult.Yes)
            {
                ExcluirBd(intValorAntigo);
            }
            ConsultarBd();
        }

        private void bndnavbtnConfirmar_Click(object sender, EventArgs e)
        {
            if (boolValorInserido)
            {
                if (MessageBox.Show("Deseja Incluir " + dtgwPreferencias.CurrentCell.EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    IncluirBd(dtgwPreferencias.CurrentCell.EditedFormattedValue.ToString());
                }
                boolValorInserido = false;
                ConsultarBd();
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar " + strValorAntigo + " Para " + dtgwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    AlterarBd(intValorAntigo, dtgwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
                }
                ConsultarBd();
            }
        }

        private void bndnavbtnPesquisa_Click(object sender, EventArgs e)
        {
            ConsultarBd(null, bndnavtxtPesquisa.Text);
        }

 // -------------------------------------FAMILIARES-------------------------------------------------------

        public void ConsultarBdFam(int? intCod = null,string strNome = null)
        {
            try // a funcao try acopla os comandos da funcao
            {
                objFamiliarVO = new FamiliaresVO();

                if (!string.IsNullOrEmpty(intCod.ToString()))
                {
                    objFamiliarVO.setCod(Convert.ToInt32(intCod));
                }
                objFamiliarVO.setNome(strNome);//seter classico ao inves do setter da ms

                objFamiliar = new Familiares();
                bndsrcFamiliares.DataSource = objFamiliar.ConsultarBd(objFamiliarVO);

                //Cria o datagrid view manualemnte para a insercao de um campo combobox
                //insercao de todas as colunas para que o datagridview nao preenche erroneamente ou fora de ordem
                dtgvvwFamiliares.Columns.Add("COD", "Codigo Familiar");
                dtgvvwFamiliares.Columns["COD"].DataPropertyName = "COD";

                dtgvvwFamiliares.Columns.Add("Nome", "Nome Familiar");
                dtgvvwFamiliares.Columns["Nome"].DataPropertyName = "Nome";

                //criacao do combobox
                DataGridViewComboBoxColumn objColunaComboxSexoSelecionavel = new DataGridViewComboBoxColumn();
                objColunaComboxSexoSelecionavel.Name = "Sexo";
                objColunaComboxSexoSelecionavel.ValueType = typeof(string);//declarando para o combo box
                objColunaComboxSexoSelecionavel.HeaderText = "Sexo Do Familiar";
                objColunaComboxSexoSelecionavel.Items.Add("MASCULINO");
                objColunaComboxSexoSelecionavel.Items.Add("FEMININO");
                objColunaComboxSexoSelecionavel.Items.Add("INDEFINIDO");
                objColunaComboxSexoSelecionavel.DataPropertyName = "Sexo";


                dtgvvwFamiliares.Columns.Add(objColunaComboxSexoSelecionavel);
                dtgvvwFamiliares.Columns["Sexo"].ValueType = typeof(string);//declarando para o dtgvw

                // insercao do restante dos campos apos o combobox
                dtgvvwFamiliares.Columns.Add("Idade", "Idade Do Familiar");
                dtgvvwFamiliares.Columns["Idade"].DataPropertyName = "Idade";

                dtgvvwFamiliares.Columns.Add("GanhoTotalMensal", "GanhoTotalMensal Do Familiar");
                dtgvvwFamiliares.Columns["GanhoTotalMensal"].DataPropertyName = "GanhoTotalMensal";

                dtgvvwFamiliares.Columns.Add("GastoTotalMensal", "GastoTotalMensal Do Familiar");
                dtgvvwFamiliares.Columns["GastoTotalMensal"].DataPropertyName = "GastoTotalMensal";

                dtgvvwFamiliares.Columns.Add("Observacao", "Observaçoes Do Familiar");
                dtgvvwFamiliares.Columns["Observacao"].DataPropertyName = "Observacao";

                //link do datagriview ao bndsrc que por sua vez ja pegou a tabela familiares
                dtgvvwFamiliares.DataSource = bndsrcFamiliares;

                // protecao contra criacao de linhas pelo usuario
                dtgvvwFamiliares.AllowUserToAddRows = false;

                //Povoamento do cmbbx de fm que lider o grd da aba de pref de fm (3er Tabela)
                //Inicializaçao do cmbbx
                //cmbbxPrefFam.DataSource = null;
                //cmbbxPrefFam.Items.Clear();

                cmbbxPrefFam.DataSource = bndsrcFamiliares.DataSource;
                cmbbxPrefFam.DisplayMember = "Nome";//parametro exhibido, texto mostrado
                cmbbxPrefFam.ValueMember = "COD";//parametro detras do texto, codigo por detras do texto
                //cmbbxPrefFam.SelectedValue = - 1;
                cmbbxPrefFam.SelectedIndex = Convert.ToInt32(intCod > 0 ? intCod - 1 : 0);
                //
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problemas na Consulta De Familiares " + ex.Message);
            }
        }

        public void ExcluirBdFam(int intCod)
        {
            try
            {
                objFamiliarVO = new FamiliaresVO();
                objFamiliarVO.setCod(intCod);
                objFamiliar = new Familiares();

                if (objFamiliar.ExcluirBd(objFamiliarVO))
                {
                    MessageBox.Show("Exclusao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Exclusao!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        public void IncluirBdFam(string strNome = null, 
                                 string strSexo = null, 
                                 int? intIdade = null,
                                 double? dbGanho= null,
                                 double? dbGasto = null,
                                 string strObs = null)
        {
            try
            {
                objFamiliarVO = new FamiliaresVO();

                objFamiliarVO.Nome = strNome;
                objFamiliarVO.Sexo = strSexo;
           
                if (intIdade != null)
                {
                    objFamiliarVO.setIdade(Convert.ToInt32(intIdade));
                }

                objFamiliarVO.Ganho = Convert.ToDouble(dbGanho == null ? 0 : dbGanho); // if inline
                objFamiliarVO.Gasto = Convert.ToDouble(dbGasto == null ? 0 : dbGasto); // necessita atribuir o zero
                objFamiliarVO.Obs =  strObs;

                objFamiliar = new Familiares();
                if (objFamiliar.IncluirBd(objFamiliarVO))
                {
                    MessageBox.Show("Inclusao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Inclusao!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }

        public void AlterarBdFam(int intCod,
                                 string strNome,
                                 string strSexo,
                                 int intIdade,
                                 double dbGanho,
                                 double dbGasto,
                                 string strObs)
        {
            try
            {
                objFamiliarVO = new FamiliaresVO(intCod, strNome, strSexo, intIdade, dbGanho, dbGasto, strObs);
                objFamiliar = new Familiares();

                if (objFamiliar.AlterarBd(objFamiliarVO))
                {
                    MessageBox.Show("Alteracao Efetuada!");
                }
                else
                {
                    MessageBox.Show("Erro na Alteracao!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ops Panguada a Vista em : " + ex.Message);
            }
        }


        private void bindingNavigatorAddNewItem1_Click(object sender, EventArgs e)
        {
            boolValorInseridoFam = true;
        }

        private void bindingNavigatorDeleteItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir " + strValorAntigo, "Aviso", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                ExcluirBdFam(intValorAntigo);
            }
            ConsultarBdFam();
        }

        private void bndnavbtnConfirmarFam_Click(object sender, EventArgs e)
        {
            if (boolValorInseridoFam)
            {
                if (MessageBox.Show("Deseja Incluir " + dtgvvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    IncluirBdFam(dtgvvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                 dtgvvwFamiliares.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                 Convert.ToInt32(dtgvvwFamiliares.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                 Convert.ToDouble(dtgvvwFamiliares.CurrentRow.Cells["GanhoTotalMensal"].EditedFormattedValue.ToString()),
                                 Convert.ToDouble(dtgvvwFamiliares.CurrentRow.Cells["GastoTotalMensal"].EditedFormattedValue.ToString()),
                                 dtgvvwFamiliares.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                }
                boolValorInseridoFam = false;
                ConsultarBdFam();
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar " + strValorAntigo + " Para " + dtgvvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
                {
                    AlterarBdFam(intValorAntigo,
                                 dtgvvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                 dtgvvwFamiliares.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                 Convert.ToInt32(dtgvvwFamiliares.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                 Convert.ToDouble(dtgvvwFamiliares.CurrentRow.Cells["GanhoTotalMensal"].EditedFormattedValue.ToString()),
                                 Convert.ToDouble(dtgvvwFamiliares.CurrentRow.Cells["GastoTotalMensal"].EditedFormattedValue.ToString()),
                                 dtgvvwFamiliares.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                }
                ConsultarBdFam();
            }
        }

        private void bndnavbtnPesquisaFam_Click(object sender, EventArgs e)
        {
            ConsultarBdFam(null, bndnavtxtPesquisaFam.Text);
        }

        private void dtgvvwFamiliares_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            strValorAntigo = dtgvvwFamiliares.CurrentRow.Cells["Nome"].Value.ToString();
            if (!string.IsNullOrEmpty(dtgvvwFamiliares.CurrentRow.Cells["COD"].Value.ToString()))
            {
                intValorAntigo = Convert.ToInt32(dtgvvwFamiliares.CurrentRow.Cells["COD"].Value.ToString());
            }
            dtgvvwFamiliares.CurrentRow.Cells["COD"].ReadOnly = true;
        }

// -------------------------------------PREFERENCIAS DE FAMILIAR ------------------------------------------

        private void cmbbxPrefFam_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(((ComboBox)sender).SelectedText))
            {
                ((ComboBox)sender).Text = ((ComboBox)sender).Text.Trim();
                dtgvvwPrefamRefresh();
            }
        }

        private void cmbbxPrefFam_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(((ComboBox)sender).ValueMember.ToString()))
            {
                dtgvvwPrefamRefresh();
            }
        }

        private void bndnavbtnPesquisar_Click(object sender, EventArgs e)
        {
            // defende codigo contra insercao de dados no cmbbxPrefFam
            try
            {
                if (!string.IsNullOrEmpty(bndnavcmbbxPrefFam.Text.Trim()))
                {
                    ConsultarPrefFam(Convert.ToInt32(cmbbxPrefFam.SelectedValue.ToString()),
                                    Convert.ToInt32(bndnavcmbbxPrefFam.Text.Substring(0, bndnavcmbbxPrefFam.Text.IndexOf("-"))),
                                    cmbbxPrefFam.Text,
                                    bndnavcmbbxPrefFam.Text.Substring(bndnavcmbbxPrefFam.Text.IndexOf("-") + 1));
                }
                else
                {
                    dtgvvwPrefamRefresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lesado.. isto e um COMBOX, Escolha uma opcao.. : " + ex.Message);
            }
        }

        private void dtgwPreferencias_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewComboBoxColumn) //programacao do combobox
            {
                ((ComboBox)e.Control).DropDownStyle = ComboBoxStyle.DropDown; //estilo de combox
                ((ComboBox)e.Control).AutoCompleteSource =  AutoCompleteSource.ListItems; // como aparece o primeiro do combobox
                ((ComboBox)e.Control).AutoCompleteMode = AutoCompleteMode.Suggest; // sugestoes na 1 lst
            }
        }
        
        private void dtgvvwPrefamRefresh() // o ID É nulo pq quero que retorne todas as preferencias do familiar
        {
            try
            {
                ConsultarPrefFam(Convert.ToInt32(cmbbxPrefFam.SelectedValue.ToString()), null, 
                                 cmbbxPrefFam.Text, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Consulta na Preferencia De Familiar ==> " + ex.Message);
            }
        }
        
        public void ConsultarPrefFam(int? intCod = null, 
                                     int? intiD = null, 
                                     string strNome = null, 
                                     string strDescricao = null)
        {
            try
            {
                objPrefFamBLL = new PrefFamBLL();
                objPrefFamVO = new PrefFamVO(); 
                objFamiliarVO = new FamiliaresVO();
                objFamiliarVO.COD = Convert.ToInt32(intCod == null ? 0 : intCod);
                objFamiliarVO.Nome = strNome;
                objPreferenciaVO = new PreferenciaVO();
                objPreferenciaVO.ID = Convert.ToInt32(intiD == null ? 0 : intiD);
                objPreferenciaVO.Descricao = strDescricao;
                objPrefFamVO.FamiliarVO = objFamiliarVO;
                objPrefFamVO.PreferenciaVO = objPreferenciaVO;

                bndsrcPrefFam.DataSource = objPrefFamBLL.ConsultarBd(objPrefFamVO);
                dtgvvwPrefFam.Columns.Clear();
                dtgvvwPrefFam.DataSource = null;
                dtgvvwPrefFam.AllowUserToAddRows = false;

                Preferencia objPreferencias = new Preferencia();
                bndsrcPreferenciaLookUp.DataSource = objPreferencia.ConsultarBd(objPreferenciaVO);

                DataGridViewComboBoxColumn objColumnaComboBoxFamiliarLookUp = new DataGridViewComboBoxColumn();
                objColumnaComboBoxFamiliarLookUp.DataSource = bndsrcFamiliares.DataSource;
                objColumnaComboBoxFamiliarLookUp.Name = "COD";
                objColumnaComboBoxFamiliarLookUp.ValueType = typeof(int);
                objColumnaComboBoxFamiliarLookUp.ValueMember = "COD";
                objColumnaComboBoxFamiliarLookUp.DisplayMember = "Nome";
                objColumnaComboBoxFamiliarLookUp.HeaderText = "Identificaçao do Familiar";
                objColumnaComboBoxFamiliarLookUp.DataPropertyName = "COD";
                dtgvvwPrefFam.Columns.Add(objColumnaComboBoxFamiliarLookUp);
                dtgvvwPrefFam.Columns["COD"].ValueType = typeof(int);
                dtgvvwPrefFam.Columns["COD"].DataPropertyName = "COD";

                DataGridViewComboBoxColumn objColumnaComboBoxPreferenciaLookUp = new DataGridViewComboBoxColumn();
                objColumnaComboBoxPreferenciaLookUp.DataSource = bndsrcPreferenciaLookUp.DataSource;
                objColumnaComboBoxPreferenciaLookUp.Name = "ID";
                objColumnaComboBoxPreferenciaLookUp.ValueType = typeof(int);
                objColumnaComboBoxPreferenciaLookUp.ValueMember = "ID";
                objColumnaComboBoxPreferenciaLookUp.DisplayMember = "Descricao";
                objColumnaComboBoxPreferenciaLookUp.HeaderText = "Descriçao da Preferencia do Familiar";
                objColumnaComboBoxPreferenciaLookUp.DataPropertyName = "ID";
                dtgvvwPrefFam.Columns.Add(objColumnaComboBoxPreferenciaLookUp);
                dtgvvwPrefFam.Columns["ID"].ValueType = typeof(int);
                dtgvvwPrefFam.Columns["ID"].DataPropertyName= "ID";

                dtgvvwPrefFam.Columns.Add("Intensidade", "Intensidades da Preferencia Do Familiar");
                dtgvvwPrefFam.Columns["Intensidade"].DataPropertyName = "Intensidade";

                dtgvvwPrefFam.Columns.Add("Observacao", "Observaçoes da Preferencia Do Familiar");
                dtgvvwPrefFam.Columns["Observacao"].DataPropertyName = "Observacao";

                dtgvvwPrefFam.DataSource = bndsrcPrefFam;
                
                //Limpa e Inicializa o cmbbx do bndnav
                bndnavcmbbxPrefFam.Items.Clear();
                //resolver possibilidade de inserir nada
                bndnavcmbbxPrefFam.Items.Add("0- ");

                foreach (DataRow objPreferenciaLinha in ((DataTable)bndsrcPreferencias.DataSource).Rows)
                {
                    bndnavcmbbxPrefFam.Items.Add(
                        objPreferenciaLinha["ID"].ToString() + "-" +
                        objPreferenciaLinha["Descricao"].ToString());
                }
                
                // TODO INIBIR O PREENCHIMENTO MANUAL DO COMBO BOX bndnavcmbbxPrefFam.Text = false;
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falhas na Consulta De Preferencias Dos Familiares" + ex.Message);
            }
        }

        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            boolInserirPrefFam = true;
            dtgvvwPrefFam.CurrentRow.Cells["COD"].Selected = false;
            dtgvvwPrefFam.CurrentRow.Cells["COD"].ReadOnly = true;
            dtgvvwPrefFam.CurrentRow.Cells["ID"].Selected = true;
            dtgvvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = false;
        }

        private void IncluirBdPrefFam(int intCod,string strNome,int intId,string strDescricao, float fltIntensidade, string strObsevacao)
        {
            try
            {
                objPrefFamVO = new PrefFamVO();
                objPrefFamVO.FamiliarVO = new FamiliaresVO();
                objPrefFamVO.FamiliarVO.COD = intCod;
                objPrefFamVO.FamiliarVO.Nome = strNome;
                objPrefFamVO.PreferenciaVO = new PreferenciaVO(intId, strDescricao);
                objPrefFamVO.Intensidade = fltIntensidade;
                objPrefFamVO.Observacao = strObsevacao;

                objPrefFamBLL = new PrefFamBLL();

                if (objPrefFamBLL.IncluirBd(objPrefFamVO))
                {
                    MessageBox.Show("Preferencia de Familiar Incluida com Sucesso!");
                }
                else
                {
                    MessageBox.Show("Problemas de Inclusao de Preferencia de Familiar!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problemas de Inclusao de Preferencia de Familiar :" + ex.Message);
            }
        }
        private void bindingNavigatorDeleteItem2_Click(object sender, EventArgs e)
        {
            ExcluirPrefFam(IntPrefFamCod, strValorAntigoPrefFamNome, intPrefFamId, strValorAntigoDescricao);
            dtgvvwPrefamRefresh();
        }
        public void ExcluirPrefFam(int intCodPrefFam,string strNome, int intIdPrefFam, string strDescricao)
        {
            try
            {
                objPrefFamVO = new PrefFamVO();
                objPrefFamVO.FamiliarVO = new FamiliaresVO();
                objPrefFamVO.FamiliarVO.COD = intCodPrefFam;
                objPrefFamVO.FamiliarVO.Nome = strNome;
                objPrefFamVO.PreferenciaVO = new PreferenciaVO(intIdPrefFam,strDescricao);
                objPrefFamBLL = new PrefFamBLL();

                if (objPrefFamBLL.ExcluirBd(objPrefFamVO))
                {
                    MessageBox.Show("Preferencia Excluida !");
                }
                else
                {
                    MessageBox.Show("Erro  na Preferencia a Excluir!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro  na Preferencia a Excluir : " + ex.Message);
            }
        }

        private void bndnavConfirmar_Click(object sender, EventArgs e)
        {
            if (boolInserirPrefFam)
            {
                if (MessageBox.Show("Deseja Inserir " + cmbbxPrefFam.Text + "para a Preferencia " + dtgvvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(), "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    IncluirBdPrefFam(Convert.ToInt32(dtgvvwPrefFam.CurrentRow.Cells["COD"].Value.ToString()),
                                dtgvvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString(),
                                Convert.ToInt32(dtgvvwPrefFam.CurrentRow.Cells["ID"].Value.ToString()),
                                dtgvvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(),
                                Convert.ToSingle(dtgvvwPrefFam.CurrentRow.Cells["Intensidade"].EditedFormattedValue.ToString()),
                                dtgvvwPrefFam.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                }
                boolInserirPrefFam = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar a preferencia do familiar " + strValorAntigoPrefFamNome + "com a sua preferencia " + strValorAntigoDescricao + "para a Preferencia ", "Aviso", MessageBoxButtons.OKCancel) == System.Windows.Forms.DialogResult.OK)
                {
                    AlterarBdPrefFam(IntPrefFamCod,
                                    strValorAntigoPrefFamNome,       
                                    intPrefFamId,
                                    strValorAntigoDescricao,
                                    Convert.ToSingle(dtgvvwPrefFam.CurrentRow.Cells["Intensidade"].EditedFormattedValue.ToString()),
                                    dtgvvwPrefFam.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                }
            }
            dtgvvwPrefamRefresh();
        }

        private void dtgvvwPrefFam_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgvvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString()) &&
                !string.IsNullOrEmpty(dtgvvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString()))
            {
                IntPrefFamCod = Convert.ToInt32(dtgvvwPrefFam.CurrentRow.Cells["COD"].Value.ToString());
                strValorAntigoPrefFamNome = dtgvvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString();
                intPrefFamId = Convert.ToInt32(dtgvvwPrefFam.CurrentRow.Cells["ID"].Value.ToString());
                strValorAntigoDescricao = dtgvvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString();
                dtgvvwPrefFam.CurrentRow.Cells["COD"].ReadOnly = true;
                dtgvvwPrefFam.CurrentRow.Cells["COD"].Selected = false;
                dtgvvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = true;
                dtgvvwPrefFam.CurrentRow.Cells["ID"].Selected = false;
            }
            else
            {
                // povoamento do codgio liderado pelo combobox
                dtgvvwPrefFam.CurrentRow.Cells["COD"].Value = cmbbxPrefFam.SelectedValue;
                dtgvvwPrefFam.CurrentRow.Cells["COD"].ReadOnly = true;
                dtgvvwPrefFam.CurrentRow.Cells["COD"].Selected = false;
                dtgvvwPrefFam.CurrentRow.Cells["ID"].Selected = true;
                dtgvvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = false;
            }
        }

        private void AlterarBdPrefFam(int intCod, string strNome, int intId, string strDescricao, float fltIntensidade, string strObsevacao = null)
        {
            try
            {
                objPrefFamVO = new PrefFamVO();
                objPrefFamVO.FamiliarVO = new FamiliaresVO();
                objPrefFamVO.FamiliarVO.COD = intCod;
                objPrefFamVO.FamiliarVO.Nome = strNome;
                objPrefFamVO.PreferenciaVO = new PreferenciaVO(intId, strDescricao);
                objPrefFamVO.Intensidade = fltIntensidade;
                objPrefFamVO.Observacao = strObsevacao;

                objPrefFamBLL = new PrefFamBLL();

                if (objPrefFamBLL.AlterarBd(objPrefFamVO))
                {
                    MessageBox.Show("Preferencia de Familiar Alterada com Sucesso!");
                }
                else
                {
                    MessageBox.Show("Problemas de Alteracao de Preferencia de Familiar!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problemas de Alteracao de Preferencia de Familiar :" + ex.Message);
            }
        }

        // <-------------------- automação do excel pelo grid -------------------->
        private void bndnavbtnAutExcel_Click(object sender, EventArgs e)
        {
            AutomacaoDeExcelPeloGrid(dtgwPreferencias);
        }
        private void bndnavbtnAutExcFam_Click(object sender, EventArgs e)
        {
            AutomacaoDeExcelPeloGrid(dtgvvwFamiliares);
        }
        private void bndnavbtnAutExcPreffam_Click(object sender, EventArgs e)
        {
            AutomacaoDeExcelPeloGrid(dtgvvwPrefFam);

        }
        public void AutomacaoDeExcelPeloGrid(DataGridView dtgdvwdeTrabalho)
        {
            //cria a automacao excel
            objExcelApplicacao = new Excel.Application(); // sem underline
            // torna visivel ou invisivel a aplicacao excel
            objExcelApplicacao.Visible = true;
            // cria o arquico excel utilizando esta automacao
            objExcelArquivo = objExcelApplicacao.Workbooks.Add();

            //associa uma planilha dentro do arquivo excel criado
            objExcelPlanilha = objExcelArquivo.Worksheets[1];

            //cria os objetos celulas com range a partir a planilha criada
            int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;

            objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
            objExcelPlanilhaCabecalho = objExcelPlanilha.Cells[intLinhaCabecalho, intColuna];

            //atribuicao de valores apra verificar o funcionamento da planilha
            //objExcelPlanilhaCabecalho.set_Value(Type.Missing, "Teste de Cabecalho do Excel");
            //objExcelPlanilhaDados.set_Value(Type.Missing, "Teste de Dados da planilha do Excel");

            //algoritmo de preenchimento da planilha excel comdois laços o mais externo para 
            //linhas primeirramente e o mais intermo pára as colunas dentro da outra 

            foreach (DataGridViewRow objLinhaGrid in dtgdvwdeTrabalho.Rows)
            {
                foreach (DataGridViewColumn objColunaGrid in dtgdvwdeTrabalho.Columns)
                {
                    if (intLinha <= 2)
                    {
                        objExcelPlanilhaCabecalho.set_Value(Type.Missing, objColunaGrid.HeaderText.ToString());
                    }

                    if (objLinhaGrid.Cells[intColuna - 1].Value != null)
                    {
                        //atribuicao da celula de dados contendo os valores do banco de dados que estao no grid
                        // para essa celula defendendo contra null no excel

                        objExcelPlanilhaDados.set_Value(Type.Missing, objLinhaGrid.Cells[intColuna - 1].Value.ToString());
                    }

                    //incremento da coluna para os objetos do excel
                    intColuna++;

                    if (intLinha <= 2)
                    {
                        objExcelPlanilhaCabecalho = objExcelPlanilha.Cells[intLinhaCabecalho, intColuna];
                    }

                    //atribuicao do objeto planilha dados para a proxima celula relativa a proxima coluna
                    objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
                }
                //incremtnto da linha para os objetos excel
                intLinha++;
                //traz o ponteiro para a primeira coluna
                intColuna = 1;

                //atribuicao do objeto planilha dados para a proxima linha na sua primeira celula
                objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
            }

            //salva o arquivo excel em um arquivo nomeado
            objExcelArquivo.SaveAs(@"C:\curso_de_programacao\AutomacaoExcelPeloGridDe" + dtgdvwdeTrabalho.Name.Substring(6) + " " + DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared);


            //fecha excel
            objExcelApplicacao.Quit();

            //avisa do termino
            MessageBox.Show("Exportação para Excel Concluida! " + dtgdvwdeTrabalho.Name.Substring(6) + " " + DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", "Exportar Excel Data Grid View", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        // <-------------------- automação do excel pela Consulta ao BD -------------------->
        private void bndnavbtnAutExcelPeloBd_Click(object sender, EventArgs e)
        {
            objPrefFamBLL= new PrefFamBLL();
            //chamada nomeada
            objPrefFamVO = new PrefFamVO();

            objFamiliarVO = new FamiliaresVO();
            objFamiliarVO.setCod(Convert.ToInt32(cmbbxPrefFam.SelectedValue.ToString()));
            objFamiliarVO.setNome(cmbbxPrefFam.Text);
            objPreferenciaVO = new PreferenciaVO();
            objPreferenciaVO.ID = 0;

            objPrefFamVO.FamiliarVO = objFamiliarVO;
            objPrefFamVO.PreferenciaVO = objPreferenciaVO;

            AutomacaoDeExcelPeloBd(objPrefFamBLL.ConsultarBd(objPrefFamVO));
                
        }
        private void bndnavbtnAutExcPrefFamBd_Click(object sender, EventArgs e)
        {
            objPrefFamBLL = new PrefFamBLL();
            //chamada nomeada
            objPrefFamVO = new PrefFamVO();

            objFamiliarVO = new FamiliaresVO();
            objFamiliarVO.setCod(Convert.ToInt32(cmbbxPrefFam.SelectedValue.ToString()));
            objFamiliarVO.setNome(cmbbxPrefFam.Text);
            objPreferenciaVO = new PreferenciaVO();
            objPreferenciaVO.ID = 0;

            objPrefFamVO.FamiliarVO = objFamiliarVO;
            objPrefFamVO.PreferenciaVO = objPreferenciaVO;

            AutomacaoDeExcelPeloBd(objPrefFamBLL.ConsultarBd(objPrefFamVO));
                
        }
        private void bndnavbtnAutExcFamBd_Click(object sender, EventArgs e)
        {
            AutomacaoDeExcelPeloBd(new Familiares().ConsultarBd(new FamiliaresVO()));
        }
        public void AutomacaoDeExcelPeloBd(DataTable dtgdvwdeTrabalho)
        {
            //cria a automacao excel
            objExcelApplicacao = new Excel.Application(); // sem underline
            // torna visivel ou invsivel a aplicacao excel
            objExcelApplicacao.Visible = true;
            // cria o arquico excel utilizando esta automacao
            objExcelArquivo = objExcelApplicacao.Workbooks.Add();

            //associa uma planilha dentro do arquivo excel criado
            objExcelPlanilha = objExcelArquivo.Worksheets[1];

            //cria os objetos celulas com range a partir a planilha criada
            int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;

            objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
            objExcelPlanilhaCabecalho = objExcelPlanilha.Cells[intLinhaCabecalho, intColuna];

            //algoritimo de preencimento da planilha da com a datatable
            foreach (DataRow objLinhaBd in dtgdvwdeTrabalho.Rows)
            {
                foreach (DataColumn objColunaBd in dtgdvwdeTrabalho.Columns)
                {
                    if (intLinha <= 2)
                    {
                        objExcelPlanilhaCabecalho.set_Value(Type.Missing, objColunaBd.ColumnName);
                    }

                    if (!string.IsNullOrEmpty(objLinhaBd[intColuna - 1].ToString()))
                    {
                        //atribuicao da celula de dados contendo os valores do banco de dados que estao no grid
                        // para essa celula defendendo contra null no excel

                        objExcelPlanilhaDados.set_Value(Type.Missing, objLinhaBd[intColuna - 1].ToString());
                    }


                    //incremento da coluna para os objetos do excel
                    intColuna++;

                    if (intLinha <= 2)
                    {
                        objExcelPlanilhaCabecalho = objExcelPlanilha.Cells[intLinhaCabecalho, intColuna];
                    }

                    //atribuicao do objeto planilha dados para a proxima celula relativa a proxima coluna
                    objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
                }
                intLinha++;
                //traz o ponteiro para a primeira coluna
                intColuna = 1;

                //atribuicao do objeto planilha dados para a proxima linha na sua primeira celula
                objExcelPlanilhaDados = objExcelPlanilha.Cells[intLinha, intColuna];
            }


            objExcelArquivo.SaveAs(@"C:\curso_de_programacao\AutomacaoExcelPeloBdDe_" + dtgdvwdeTrabalho.TableName + " " + DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx",
      Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared);


            //fecha excel
            objExcelApplicacao.Quit();

            //avisa do termino
            MessageBox.Show("Exportação para Excel do Banco de Dados Concluida! AutomacaoExcelPeloBdDe_" + dtgdvwdeTrabalho.TableName + " " + DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", "Exportar Excel Banco de Dados", MessageBoxButtons.OK, MessageBoxIcon.Hand);
        }

        // <-------------------- automação do excel pela Consulta ao BD -------------------->
        private void bndnavPrefbtnAutExcAcInt_Click(object sender, EventArgs e)
        {
            try
            {
                objPreferencia = new Preferencia();

                sfdPlanilhaInterop.ShowDialog();

                objPreferencia.GeraExcelDoAccessPorinterop(sfdPlanilhaInterop.FileName);

                MessageBox.Show("A exportação foi concluida com sucesso! ","Exportação do Acesso Interop Excel",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void bndnavbtnFamIntAcEx_Click(object sender, EventArgs e)
        {
            try
            {
                objFamiliar = new Familiares();

                sfdPlanilhaInterop.ShowDialog();

                objFamiliar.GeraExcelDoAccessPorinterop(sfdPlanilhaInterop.FileName);

                MessageBox.Show("A exportação foi concluida com sucesso! ", "Exportação do Acesso Interop Excel", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void bndnavbtnPrefFamIntAccExc_Click(object sender, EventArgs e)
        {
            try
            {
                objPrefFamBLL = new PrefFamBLL();

                sfdPlanilhaInterop.ShowDialog();

                objPrefFamBLL.GeraExcelDoAccessPorinterop(sfdPlanilhaInterop.FileName,Convert.ToInt32(cmbbxPrefFam.SelectedValue.ToString()));

                MessageBox.Show("A exportação foi concluida com sucesso! ", "Exportação do Acesso Interop Excel", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // <-------------------- automação do excel pela Consulta ao BD -------------------->
        private void bndnavPreferenciaEnviaEmail_Click(object sender, EventArgs e)
        {
            GeraEmailOutLook();

        }
        private void bnfnavbtnEnvAutFam_Click(object sender, EventArgs e)
        {
            GeraEmailOutLook();
        }
        private void bnfnavbtnEnvAutPrefFam_Click(object sender, EventArgs e)
        {
            GeraEmailOutLook();
        }        
        public void GeraEmailOutLook()
        {
            try
            {
                //Cria Automaçção de Outlook para Email - Cria objeto de Automação Outlook
                objEmailApp = new Email.Application();
                //Cria Email - Menssagem utilizando a Aplicação Outlook criada - Cria Objeto Mensagem de Outlook
                objMensagem = objEmailApp.CreateItem(Email.OlItemType.olMailItem);
                //Povoa Objeto mensagem
                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                objMensagem.SentOnBehalfOfName = "tilson2000@hotmail.com";
                
                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                //Especipico o(s) Email(s) Destinatário Principal ou Destinatários Principais das mensagens - Endereço "Para" quem é o email
                objMensagem.To = "robertosptcosta@gmail.com";

                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                //Especipico o(s) Email(s) Destinatário Em Cópia ou Destinatários Copiados das mensagens - Endereço "CC" com Cópia do email
                objMensagem.CC = "a_bsilva1@hotmail.com";

                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                //Especipico o(s) Email(s) Destinatário Em Cópia Oculta ou Destinatários Copiados Em Oculto das mensagens - Endereço "BCC" com Cópia Oculta ou BCC - Blind Copy do email
                objMensagem.BCC = "tilson2000@hotmail.com ";

                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                //Assunto do Email - Subject 
                objMensagem.Subject = "Teste de Email Automático - Teste do Treinamento de Envio de E-mail pelo Outlook Para os Avançados";

                //Especipico o Email pelo qual será enviada as mensagens - Caixa de email de origem das mensagens - Endereço "De" quem está enviando o email
                //Corpo - Mensagem "Body" do E-mail
                objMensagem.Body = "Pessoal, Bom Dia !"  +
                                    Environment.NewLine  +
                                    "Conforme combinado, segue esse testo de email para o envio automático do mesmo pelo C#. \n" +
                "(Esse é um e-mail automático, não responda)";

                
                // pergunta se deseja anexar
                if (MessageBox.Show("Deseja Anexar Arquivos ? "," Aviso ", MessageBoxButtons.YesNo,MessageBoxIcon.Question)== System.Windows.Forms.DialogResult.Yes)
	            {
		         //utilização do open file dialog
                    //Título da Janela
                    ofdEmailArquivoAnexo.Title = "Escolha os Arquivos a Serem Anexados ao E-mail";

                    //Configuro a pasta inicial de pesquisa dos arquivos a serem anexados
                    ofdEmailArquivoAnexo.InitialDirectory = @"C:\curso_de_programacao";

                    //Executa a janela de Diálogo do Open File Dialog para pegar os arquivos a serem 
                    // anexados - pode fazer seleção múltipla de arquivos - para anexar mais de um arquivo

                    //Anexo - Arquivos Anexos "Attachments" - Adicionamento ao Email
                    ofdEmailArquivoAnexo.ShowDialog();

                    //Verifico a lista de Arquivos Anexos - "ofdArqAnexosEmail.FileName" - se não está 
                    //vazia ou nula para executar o laço de anexação de arquivos a serem anexados 
                    if (!string.IsNullOrEmpty(ofdEmailArquivoAnexo.FileName))
	                {
                        //Alimentação da Matriz de Nomes de Arquivos Anexos
                        //Realizo a alimentação da Matriz / Array de Arquivos Anexos atravé DO PLURAL DO FileName, ou seja, a PROPRIEDADE FileNames- arrayEmailArqAnexo
                        arrayEmailArqAnexo = ofdEmailArquivoAnexo.FileNames;

                        //Laço para anexação de arquivos constantes na matriz de arquivos anexos - arrayEmailArqAnexo - para o email (mensagem) - objMensagem
                        for (int i = 0; i < arrayEmailArqAnexo.Length; i++)
                        {
                             //Coloco o tipo de anexo pelo valor   
                            objArquivoAnexoTipo = Email.OlAttachmentType.olByValue;

                            //Coloco a posição do arquivo a ser anexado no final do Corpo "Body" mais (uma) posição ao final
                            lngEmailArqAnexoPosicao = objMensagem.Body.Length + 1;

                            //Coloco o Nome do Anexo a ser mostrado no e-mail
                            strDisplayName = arrayEmailArqAnexo[i].ToString() + " - NovoArquivo-treinoAnexo";
                            
                            //Adiciona o Arquivo Anexo ao objeto E-mail Mensagem para ser enviado finalmente
                            objMensagem.Attachments.Add(arrayEmailArqAnexo[i],objArquivoAnexoTipo,lngEmailArqAnexoPosicao,strDisplayName);
                        }
	                }
                }

                //Envio da Mensagem com pedido de confirmação ou não (automaticamente)
             if (MessageBox.Show("Envia Email com Confirmação?", "Pedido", MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)      
              {
                     //visualiza o email e enia manualente
                     objMensagem.Display();
              }
                else
              {
                     //envia automaticamente
                     objMensagem.Send();
                     MessageBox.Show("Email Enviado com Sucesso !", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                     objEmailApp.Quit(); // pode dar problemas se não for ativado
              }
             MessageBox.Show("Termino da automação !", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro no envio automatico de Email" + ex.Message);
            }
     
        }
    }
}
