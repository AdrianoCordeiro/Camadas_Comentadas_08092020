using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Preferencia_Model_VO;

// a DAO tira da camada  de negocios (Bussiness) os acessos a dados, isola o acesso e manipulacao do BD e  
// a complexicade desse tratamento dentro de uma camada especifica criada para conter todas as especificacoes   
// e nuances doscomandos e acessos ao BD
// retirando e mitigando o impacto das alteracoes e modificacoes nos comandos instrucoes e acessos ao BD
// de impactar ou constituir (trazer) problemas para as regras de negocios da camada de negocio
namespace PreferencciaDAO_DAL 
{
    public class PreferenciaDAO : DAO_DAL
    {
        PreferenciaVO objPreferenciaVO; //cria objeto da classe modelo
        OleDbCommand objComm; // cria objeto de comando oledb
        OleDbDataAdapter objAdap; // cria adaptador para acesso desconectado
        OleDbDataReader objLeitorBd; // cria leitor de bancos para acesso conectado
        DataTable objTabela; // cria objeto de tabela para dados

        public List<string> ImportarBdConectado()
        {
            try
            {
                List<string> resultado = new List<string>();

                AbreConexao();

                objComm = new OleDbCommand(@"SELECT ID, Descricao FROM Preferencias",getConexao());

                objLeitorBd = objComm.ExecuteReader();

                while (objLeitorBd.Read())
                {
                    resultado.Add(objLeitorBd["Descricao"].ToString());
                }
                objLeitorBd.Close();

                return resultado;

            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Importar Conectado do Banco de Dados : " + ex.Message);
            }
            finally
            {
                FechaConexao();
            }
        }

        public List<string> ImportarBdDesconectado()
        {
            try
            {
                List<string> resultado = new List<string>();


                objComm = new OleDbCommand(@"SELECT ID, Descricao FROM Preferencias",getConexao());

                objAdap = new OleDbDataAdapter();
                objAdap.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdap.Fill(objTabela);

                foreach (DataRow strLinhaTabela in objTabela.Rows)
                {
                    resultado.Add(strLinhaTabela["Descricao"].ToString());
                }

                return resultado;

            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Importar Desconectado do Banco de Dados : " + ex.Message);
            }
        }

        public override DataTable ConsultarBd(Object objparPreferenciaVO)
        {
            try
            {
                // stringbuilder evita a passagem de texto, sendo padrao de mercado e boa pratica
                StringBuilder strSql = new StringBuilder();
                //Fazendo o casting (trocando um tipo de objeto por outro
                objPreferenciaVO = (PreferenciaVO)objparPreferenciaVO;

                if (objPreferenciaVO.getId()>0)
                {

        // codigoo sem parametro no SQL
        //objComm.CommandText = @"SELECT Descricao FROM Preferencias WHERE Descricao = " + objparPreferenciaVO.getId();
        // textto do SQL com parametros para evitar a falha de seguranca de injecao de SQL (SQL INJECTION)
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    strSql.Append("WHERE ID = :parId");

                    objComm = new OleDbCommand(strSql.ToString(),getConexao());
                    objComm.Parameters.AddWithValue("parId", objPreferenciaVO.getId());
                }
                else if (string.IsNullOrEmpty(objPreferenciaVO.getDescricao()))
                {
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                }
                else
                { // sai a concatenacao e poe : para mostrar o parametro passado usando o objcomm. parameters addwithvalue(so access so funciona com ele)
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    strSql.Append("WHERE Descricao = :parDescricao"); 
                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parDescricao", objPreferenciaVO.getDescricao());
                }

                objAdap = new OleDbDataAdapter();
                objAdap.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdap.Fill(objTabela);

                return objTabela;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Consultar da Camada BLL do Banco de Dados : " + ex.Message);
            }
        }

        // cria uma copia da model quando passado no parametro com o ref p(pasagem pelo conteudo)
        // o ref passa o endereco da model 
        // sobrecarga de natureza de parametros
        public override void ConsultarBd(ref Object objparPreferenciaVO)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                objPreferenciaVO = (PreferenciaVO) objparPreferenciaVO;

                if (objPreferenciaVO.getId() > 0)
                {
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    strSql.Append("WHERE ID = :parId"); 
                    // codigoo sem parametro no SQL
                    //objComm.CommandText = @"SELECT Descricao FROM Preferencias WHERE Descricao = " + objparPreferenciaVO.getId();
                    // textto do SQL com parametros para evitar a falha de seguranca de injecao de SQL (SQL INJECTION)
                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parId", objPreferenciaVO.getId());
                }
                else if (string.IsNullOrEmpty(objPreferenciaVO.getDescricao()))
                {
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                }
                else
                { // sai a concatenacao e poe : para mostrar o parametro passado usando o objcomm. parameters addwithvalue(so access so funciona com ele)
                    strSql.Append("SELECT ");
                    strSql.Append("ID ");
                    strSql.Append(",Descricao ");
                    strSql.Append("FROM ");
                    strSql.Append("Preferencias ");
                    strSql.Append("WHERE Descricao = :parDescricao");
                    // codigoo sem parametro no SQL
                    //objComm.CommandText = @"SELECT Descricao FROM Preferencias WHERE Descricao = " + objparPreferenciaVO.getId();
                    // textto do SQL com parametros para evitar a falha de seguranca de injecao de SQL (SQL INJECTION)
                    objComm = new OleDbCommand(strSql.ToString(), getConexao());
                    objComm.Parameters.AddWithValue("parDescricao", objPreferenciaVO.getDescricao());
                }

                objAdap = new OleDbDataAdapter();
                objAdap.SelectCommand = objComm;

                objTabela = new DataTable();

                objAdap.Fill(objTabela);

                foreach (DataRow drItemTabela in objTabela.Rows)
                {
                    objPreferenciaVO = new PreferenciaVO(Convert.ToInt32(drItemTabela["ID"].ToString()), drItemTabela["Descricao"].ToString());

                    objPreferenciaVO.PreferenciaVOCollection.Add(objPreferenciaVO);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Consultar da Camada BLL do Banco de Dados : " + ex.Message);
            }
        } 

        public override bool IncluirBd(Object objparPreferenciaVO)
        {
            try
            {
                // faz a criacao de varial do tipo stringbuilder
                StringBuilder strSql = new StringBuilder();
                // faz o casting objeto criado = (tipo do objeto) parametro dado
                objPreferenciaVO = (PreferenciaVO)objparPreferenciaVO;

                bool resultado = false;

                AbreConexao();

                strSql.Append("INSERT INTO ");
                strSql.Append(" Preferencias (");
                strSql.Append(" Descricao) ");
                strSql.Append(" VALUES (");
                strSql.Append(" :parDescricao)");
                objComm = new OleDbCommand(strSql.ToString(),getConexao());
                objComm.Parameters.AddWithValue("parDescricao", objPreferenciaVO.Descricao);


                if (objComm.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }
                return resultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Incluir da Camada BLL do Banco de Dados : " + ex.Message);
            }
            finally // o finally roda independente de erros que o try/catch tenha achando e nesse caso
            {       // fecha os objetos keitor r conn
                FechaConexao();
            }
        }

        public override bool ExcluirBd(Object objparPreferenciaVO)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                objPreferenciaVO = (PreferenciaVO)objparPreferenciaVO;
                bool resultado = false;

                AbreConexao();
                strSql.Append("DELETE ");
                strSql.Append("FROM ");
                strSql.Append("Preferencias ");
                strSql.Append("WHERE ID = :parId");
                objComm = new OleDbCommand(strSql.ToString(),getConexao());
                objComm.Parameters.AddWithValue("parId", objPreferenciaVO.ID);

                if (objComm.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }
                return resultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Excluir da Camada BLL do Banco de Dados : " + ex.Message);
            }
            finally // o finally roda independente de erros que o try/catch tenha achando e nesse caso
            {       // fecha os objetos keitor r conn
                FechaConexao();
            }
        }

        public override bool AlterarBd(Object objparPreferenciaVO)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();

                objPreferenciaVO = (PreferenciaVO)objparPreferenciaVO;
                bool resultado = false;

                AbreConexao();

                strSql.Append("UPDATE ");
                strSql.Append("Preferencias ");
                strSql.Append("SET ");
                strSql.Append("Descricao = :parDescricaoNovo ");
                strSql.Append("WHERE ID = :parId");
                objComm = new OleDbCommand(strSql.ToString(),getConexao());
                objComm.Parameters.AddWithValue("parDescricaoNovo", objPreferenciaVO.Descricao);
                objComm.Parameters.AddWithValue("parId", objPreferenciaVO.getId());

                if (objComm.ExecuteNonQuery() > 0)
                {
                    resultado = true;
                }
                else
                {
                    resultado = false;
                }
                return resultado;
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar da Camada BLL do Banco de Dados : " + ex.Message);
            }
            finally // o finally roda independente de erros que o try/catch tenha achando e nesse caso
            {       // fecha os objetos keitor r conn
                FechaConexao();
            }
        }
        public void  GeraExcelDoAccessPorinterop(string strnNomePlanilha)
        {
            try
            {
                AbreConexao();
                // stringbuilder evita a passagem de texto, sendo padrao de mercado e boa pratica
                StringBuilder strSql = new StringBuilder();
                //Fazendo o casting (trocando um tipo de objeto por outro

                strSql.Append("SELECT");
                strSql.Append(" ID");
                strSql.Append(" ,Descricao");
                strSql.Append(" INTO");
                strSql.Append(" [EXCEL 8.0; DATABASE=" + strnNomePlanilha + "].[EXPORT EXCEL]");
                strSql.Append(" FROM");
                strSql.Append(" Preferencias");
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
