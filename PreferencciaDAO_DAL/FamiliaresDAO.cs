using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Preferencia_Model_VO;
using System.Data;
using System.Data.OleDb;

namespace PreferencciaDAO_DAL
{
    public class FamiliaresDAO : DAO_DAL
    {
        FamiliaresVO objFamiliarVO;
        OleDbCommand objComm;
        OleDbDataAdapter objAdap;
        DataTable objTabela;


        public override DataTable ConsultarBd(Object objparFamiliarVO)
        {
            try
            {
                objFamiliarVO = (FamiliaresVO)objparFamiliarVO;
                StringBuilder strSql = new StringBuilder();

                if (objFamiliarVO.getCod() > 0)
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");
                    strSql.Append(" WHERE COD = :parCod");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parCod", objFamiliarVO.getCod());
                }
                else if (string.IsNullOrEmpty(objFamiliarVO.getNome()))
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                }
                else
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");
                    strSql.Append(" WHERE Nome = :parNome");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parNome", objFamiliarVO.getNome());
                }
                objAdap = new OleDbDataAdapter();
                objAdap.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdap.Fill(objTabela);

                return objTabela;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Consultar do Banco de Dados" + ex.Message);
            }
        }

        public override void ConsultarBd(ref Object objparFamiliarVO)
        {
            try
            {
                objFamiliarVO = (FamiliaresVO)objparFamiliarVO;
                StringBuilder strSql = new StringBuilder();

                if (objFamiliarVO.getCod() > 0)
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");
                    strSql.Append(" WHERE COD = :parCod");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parCod", objFamiliarVO.getCod());
                }
                else if (string.IsNullOrEmpty(objFamiliarVO.getNome()))
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                }
                else
                {
                    strSql.Append(" SELECT");
                    strSql.Append(" COD ");
                    strSql.Append(" ,Nome ");
                    strSql.Append(" ,Sexo ");
                    strSql.Append(" ,Idade ");
                    strSql.Append(" ,GanhoTotalMensal ");
                    strSql.Append(" ,GastoTotalMensal ");
                    strSql.Append(" ,Observacao ");
                    strSql.Append(" FROM");
                    strSql.Append(" Familiares");
                    strSql.Append(" WHERE Nome = :parNome");

                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parNome", objFamiliarVO.getNome());

                }
                objAdap = new OleDbDataAdapter();
                objAdap.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdap.Fill(objTabela);

                foreach (DataRow drItemTabela in objTabela.Rows)
                {
                    objFamiliarVO = new FamiliaresVO(
                        Convert.ToInt32(drItemTabela["COD"].ToString()),
                        drItemTabela["Nome"].ToString(),
                        drItemTabela["Sexo"].ToString(),
                        Convert.ToInt32(drItemTabela["Idade"].ToString()),
                        Convert.ToDouble( drItemTabela["GanhoTotalMensal"].ToString()),
                        Convert.ToDouble( drItemTabela["GastoTotalMensal"].ToString()),
                        drItemTabela["Observacao"].ToString());

                    objFamiliarVO.objFamiliaresVOCollection.Add(objFamiliarVO);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Consultar REF do Banco de Dados" + ex.Message);
            }
        }

        public override bool IncluirBd(Object objparFamiliarVO)
        {
            try
            {
                objFamiliarVO = (FamiliaresVO)objparFamiliarVO;

                StringBuilder strSql = new StringBuilder();
                
                bool lsResultado = false;

                AbreConexao();

                strSql.Append(" INSERT INTO");
                strSql.Append(" Familiares(");
                strSql.Append(" Nome");
                strSql.Append(" ,Sexo");
                strSql.Append(" ,Idade");
                strSql.Append(" ,GanhoTotalMensal");
                strSql.Append(" ,GastoTotalMensal");
                strSql.Append(" ,Observacao)");
                strSql.Append(" VALUES(");
                strSql.Append(" :parNome");
                strSql.Append(" ,:parSexo");
                strSql.Append(" ,:parIdade");
                strSql.Append(" ,:parGanho");
                strSql.Append(" ,:parGasto");
                strSql.Append(" ,:parObs)");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parNome", objFamiliarVO.getNome());
                objComm.Parameters.AddWithValue("parSexo", objFamiliarVO.getSexo());
                objComm.Parameters.AddWithValue("parIdade", objFamiliarVO.getIdade());
                objComm.Parameters.AddWithValue("parGanho", objFamiliarVO.getGanhoTotalMensal());
                objComm.Parameters.AddWithValue("parGasto", objFamiliarVO.getGastoTotalMensal());
                objComm.Parameters.AddWithValue("parObs", objFamiliarVO.getObservacao());

                if (objComm.ExecuteNonQuery()>0)
                {
                    lsResultado = true;
                }
                else
                {
                    lsResultado = false;
                }
                return lsResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Incluir do Banco de Dados" + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public override bool ExcluirBd(Object objparFamiliarVO)
        {
            try
            {
                objFamiliarVO = (FamiliaresVO)objparFamiliarVO;
                StringBuilder strSql = new StringBuilder();
                bool lsResultado = false;

                AbreConexao();

                strSql.Append(" DELETE");
                strSql.Append(" FROM");
                strSql.Append(" Familiares");
                strSql.Append(" WHERE COD = :parCod");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                objComm.Parameters.AddWithValue("parCod", objFamiliarVO.getCod());

                if (objComm.ExecuteNonQuery()>0)
                {
                    lsResultado = true;
                }
                else
                {
                    lsResultado = false;
                }
                return lsResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Excluir do Banco de Dados" + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public override bool AlterarBd(Object objparFamiliarVO)
        {
            try
            {
                objFamiliarVO = (FamiliaresVO)objparFamiliarVO;
                StringBuilder strSql = new StringBuilder();

                bool lsResultado = false;

                AbreConexao();

                strSql.Append(" UPDATE");
                strSql.Append(" Familiares");
                strSql.Append(" SET");
                strSql.Append(" Nome = :parNome");
                strSql.Append(" ,Sexo = :parSexo");
                strSql.Append(" ,Idade = :parIdade");
                strSql.Append(" ,GanhoTotalMensal = :parGanho");
                strSql.Append(" ,GastoTotalMensal = :parGasto");
                strSql.Append(" ,Observacao = :parObs");
                strSql.Append(" WHERE COD = :parCod");

                objComm = new OleDbCommand(strSql.ToString(), getConexao());

                objComm.Parameters.AddWithValue("parNome", objFamiliarVO.getNome());
                objComm.Parameters.AddWithValue("parSexo", objFamiliarVO.getSexo());
                objComm.Parameters.AddWithValue("parIdade", objFamiliarVO.getIdade());
                objComm.Parameters.AddWithValue("parGanho", objFamiliarVO.getGanhoTotalMensal());
                objComm.Parameters.AddWithValue("parGasto", objFamiliarVO.getGastoTotalMensal());
                objComm.Parameters.AddWithValue("parObs", objFamiliarVO.getObservacao());
                objComm.Parameters.AddWithValue("parCod", objFamiliarVO.getCod());


                if (objComm.ExecuteNonQuery() > 0)
                {
                    lsResultado = true;
                }
                else
                {
                    lsResultado = false;
                }
                return lsResultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar do Banco de Dados" + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }
        public void GeraExcelDoAccessPorinterop(string strnNomePlanilha)
        {
            try
            {
                AbreConexao();
                // stringbuilder evita a passagem de texto, sendo padrao de mercado e boa pratica
                StringBuilder strSql = new StringBuilder();
                //Fazendo o casting (trocando um tipo de objeto por outro

                strSql.Append("SELECT");
                strSql.Append(" COD");
                strSql.Append(" ,Nome");
                strSql.Append(" ,Sexo");
                strSql.Append(" ,Idade");
                strSql.Append(" ,GanhoTotalMensal");
                strSql.Append(" ,GastoTotalMensal");
                strSql.Append(" ,Observacao");
                strSql.Append(" INTO");
                strSql.Append(" [EXCEL 8.0; DATABASE=" + strnNomePlanilha + "].[EXPORT EXCEL]");
                strSql.Append(" FROM");
                strSql.Append(" Familiares");
                objComm = new OleDbCommand(strSql.ToString(), getConexao());
                //objComm.Parameters.AddWithValue("parPlanilha", strnNomePlanilha);

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
