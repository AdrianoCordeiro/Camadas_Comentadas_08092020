using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Preferencia_Model_VO;

namespace PreferencciaDAO_DAL
{
    public class PrefFamDAO : DAO_DAL
    {
        OleDbCommand objComm;
        //OleDbDataReader objLeitor;
        OleDbDataAdapter objAdapt;
        DataTable objTabela;
        PrefFamVO objPrefFamVO;

        public override DataTable ConsultarBd(Object objVo_VO)
        {
            try
            {
                //executar o casting de um tipo objeto para um tipo objeto preffamvo
                objPrefFamVO = (PrefFamVO)objVo_VO; 
                StringBuilder strSql = new StringBuilder();

                // iniciando a string de comando sql
                strSql.Append("SELECT");
                strSql.Append(" COD");
                strSql.Append(" ,ID");
                strSql.Append(" ,Intensidade");
                strSql.Append(" ,Observacao");
                strSql.Append(" FROM");
                strSql.Append(" PrefFam");

                //preparando a conexao para entrada dos WHEREs
                objComm = new OleDbCommand();
                objComm.Connection = getConexao();

                //Iniciando ninho de if de verificacao do cod e do id
                if (objPrefFamVO.FamiliarVO.COD > 0 || objPrefFamVO.PreferenciaVO.ID > 0)
                {
                    //Verifica se o COD e maior que zero
                    if (objPrefFamVO.FamiliarVO.COD>0)
                    {
                        //verifica se o id e maior que zero
                        //caso as duas condicoes estejam ok preenche o where com cod e id
                        if (objPrefFamVO.PreferenciaVO.ID>0)
                        {
                            strSql.Append(" WHERE");
                            strSql.Append(" COD = :parCod");
                            strSql.Append(" AND");
                            strSql.Append(" ID = :parId");

                            objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                            objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());
                        }
                        //caso o id esteja zerado pega so o COD
                        else
                        {
                            strSql.Append(" WHERE");
                            strSql.Append(" COD = :parCod");

                            objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                        }
                    }
                    //caso o cod esteja zerado pega so o id
                    else
                    {
                        strSql.Append(" WHERE");
                        strSql.Append(" ID = :parId");

                        objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());
                    }
                }
                objComm.CommandText = strSql.ToString();

                objAdapt = new OleDbDataAdapter();
                objAdapt.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdapt.Fill(objTabela);

                return objTabela;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no COnsultar (PrefFam) do Banco de Dados " + ex.Message);
            }
        }

        public override void ConsultarBd(ref Object objVo_VO)
        {
            try
            {
                //executar o casting de um tipo objeto para um tipo objeto preffamvo
                objPrefFamVO = (PrefFamVO)objVo_VO;
                StringBuilder strSql = new StringBuilder();

                // iniciando a string de comando sql
                strSql.Append("SELECT");
                strSql.Append(" COD");
                strSql.Append(" ,ID");
                strSql.Append(" ,Intensidade");
                strSql.Append(" ,Observacao");
                strSql.Append(" FROM");
                strSql.Append(" PrefFam");

                //preparando a conexao para entrada dos WHEREs
                objComm = new OleDbCommand();
                objComm.Connection = getConexao();

                //Iniciando ninho de if de verificacao do cod e do id
                if (objPrefFamVO.FamiliarVO.COD > 0 || objPrefFamVO.PreferenciaVO.ID > 0)
                {
                    //Verifica se o COD e maior que zero
                    if (objPrefFamVO.FamiliarVO.COD > 0)
                    {
                        //verifica se o id e maior que zero
                        //caso as duas condicoes estejam ok preenche o where com cod e id
                        if (objPrefFamVO.PreferenciaVO.ID > 0)
                        {
                            strSql.Append(" WHERE");
                            strSql.Append(" COD = :parCod");
                            strSql.Append(" AND");
                            strSql.Append(" ID = :parId");

                            objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                            objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());
                        }
                        //caso o id esteja zerado pega so o COD
                        else
                        {
                            strSql.Append(" WHERE");
                            strSql.Append(" COD = :parCod");

                            objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                        }
                    }
                    //caso o cod esteja zerado pega so o id
                    else
                    {
                        strSql.Append(" WHERE");
                        strSql.Append(" ID = :parId");

                        objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());
                    }
                }
                objComm.CommandText = strSql.ToString();

                objAdapt = new OleDbDataAdapter();
                objAdapt.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdapt.Fill(objTabela);

                foreach (DataRow drItemTabela in objTabela.Rows)
                {
                    //instaciar o objfamiliar atributo familiar vo
                    objPrefFamVO.FamiliarVO = new FamiliaresVO();
                    // converter o cod do familiar vo
                    objPrefFamVO.FamiliarVO.COD = Convert.ToInt32(drItemTabela["COD"].ToString());

                    //instaciar o familiares DAO e fazer o casting
                    FamiliaresDAO objFamiliarDAO = new FamiliaresDAO();
                    Object objparObjetoFamiliarVO = (Object)objPrefFamVO.FamiliarVO;
                    //sempre que o parametro é referenciado(endereço de memoria TIPO object) 
                    //é necessario realizar o object
                    objFamiliarDAO.ConsultarBd(ref objparObjetoFamiliarVO);

                    //chama a collection da model
                    objPrefFamVO.FamiliarVO = objPrefFamVO.FamiliarVO.objFamiliaresVOCollection.First<FamiliaresVO>();

                    objPrefFamVO.PreferenciaVO = new PreferenciaVO();

                    PreferenciaDAO objPreferenciaDAO = new PreferenciaDAO();
                    Object objparObjetoPreferenciaVO = (Object)objPrefFamVO.PreferenciaVO;
                    objPreferenciaDAO.ConsultarBd(ref objparObjetoPreferenciaVO);

                    objPrefFamVO.PreferenciaVO = objPrefFamVO.PreferenciaVO.PreferenciaVOCollection.First<PreferenciaVO>();
                    objPrefFamVO.PreferenciaVO.ID = Convert.ToInt32(drItemTabela["ID"].ToString());
                    objPrefFamVO.Intensidade = Convert.ToSingle(drItemTabela["Intensidade"].ToString());
                    objPrefFamVO.Observacao = drItemTabela["Observacao"].ToString();

                    objPrefFamVO.PrefFamCollection.Add(objPrefFamVO);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no COnsultar (PrefFam) do Banco de Dados " + ex.Message);
            }
        }

        public override bool IncluirBd(Object objvo_VO)
        {
            try
            {
                AbreConexao();
                bool boolResultado = false;
                objPrefFamVO = (PrefFamVO)objvo_VO;
                StringBuilder strSql = new StringBuilder();

                strSql.Append("INSERT INTO");
                strSql.Append(" PrefFam");
                strSql.Append(" (");
                strSql.Append(" COD");
                strSql.Append(" ,ID");
                strSql.Append(" ,Intensidade");
                strSql.Append(" ,Observacao");
                strSql.Append(" ) VALUES (");
                strSql.Append(" :parCod");
                strSql.Append(" ,:parId");
                strSql.Append(" ,:parInt");
                strSql.Append(" ,:parObs");
                strSql.Append(" )");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());
                objComm.Parameters.AddWithValue("parInt", objPrefFamVO.getIntensiadade());
                objComm.Parameters.AddWithValue("parObs", objPrefFamVO.getObservacao());

                if (objComm.ExecuteNonQuery() > 0)
                {
                    boolResultado = true;
                }
                else
                {
                    boolResultado = false;
                }
                return boolResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha ao Incluir (PrefFam) do Banco de Dados " + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public override bool ExcluirBd(Object objvo_VO)
        {
            try
            {
                AbreConexao();
                bool boolResultado = false;
                objPrefFamVO = (PrefFamVO)objvo_VO;
                StringBuilder strSql = new StringBuilder();

                strSql.Append("DELETE FROM");
                strSql.Append(" PrefFam");
                strSql.Append(" WHERE");
                //chave primaria sempre no WHERE ela é concantenada, elas nao podem ser atualizadas
                strSql.Append(" COD = :parCod");
                //conector Logico, pq é pelo Cod ou ID
                strSql.Append(" AND ID = :parId");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());

                if (objComm.ExecuteNonQuery() > 0)
                {
                    boolResultado = true;
                }
                else
                {
                    boolResultado = false;
                }
                return boolResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha ao Excluir (PrefFam) do Banco de Dados " + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public override bool AlterarBd(Object objvo_VO)
        {
            try
            {
                AbreConexao();
                bool boolResultado = false;
                objPrefFamVO = (PrefFamVO)objvo_VO;
                StringBuilder strSql = new StringBuilder();

                strSql.Append("UPDATE");
                strSql.Append(" PrefFam");
                strSql.Append(" SET");
                strSql.Append(" Intensidade = :parInt");
                strSql.Append(" ,Observacao = :parObs");
                strSql.Append(" WHERE");
                // Cod e ID vao depois do WHERE que smepre que se pesquisa é pela chave primaria
                strSql.Append(" COD = :parCod");
                strSql.Append(" AND ID = :parId");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parInt", objPrefFamVO.getIntensiadade());
                objComm.Parameters.AddWithValue("parObs", objPrefFamVO.getObservacao());
                objComm.Parameters.AddWithValue("parCod", objPrefFamVO.FamiliarVO.getCod());
                objComm.Parameters.AddWithValue("parId", objPrefFamVO.PreferenciaVO.getId());

                if (objComm.ExecuteNonQuery() > 0)
                {
                    boolResultado = true;
                }
                else
                {
                    boolResultado = false;
                }
                return boolResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha ao Alterar (PrefFam) do Banco de Dados " + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public void GeraExcelDoAccessPorinterop(string strnNomePlanilha, int intCod)
        {
            try
            {
                AbreConexao();
                // stringbuilder evita a passagem de texto, sendo padrao de mercado e boa pratica
                StringBuilder strSql = new StringBuilder();
                //Fazendo o casting (trocando um tipo de objeto por outro

                strSql.Append("SELECT");
                strSql.Append(" COD");
                strSql.Append(" ,ID");
                strSql.Append(" ,Intensidade");
                strSql.Append(" ,Observacao");
                strSql.Append(" INTO");
                strSql.Append(" [EXCEL 8.0; DATABASE=" + strnNomePlanilha + "].[EXPORT EXCEL]");
                strSql.Append(" FROM");
                strSql.Append(" PrefFam");
                strSql.Append(" WHERE");
                strSql.Append(" COD = :parCod");
                
                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parCod", intCod);

                if (objComm.ExecuteNonQuery() < 1)
                {
                    throw new Exception("Erro ao Gerar Exportação do Acess por Interop na planilha " + strnNomePlanilha);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Exportar do Access Interop para Excel : " + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

    }
}
