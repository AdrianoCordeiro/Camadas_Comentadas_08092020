using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace PreferencciaDAO_DAL
{// desinger pattern singleton que isola A CONEXAO E QPERMINTE QUE ELA SEJA REUTILIZADO UTILIZANDO OS 
//RECURSOS COMPUTACIONAIS PELO REUSO DE AREA DE MEMORIA E PONTENTEIRODAS DAS MAQUINAS CLIENTES E DAS 
 //MAQUINAS SERVIDORES
    public class DB_DAO
    {
        private static OleDbConnection objConn;

        public static OleDbConnection getConexao()
        {
            if (objConn == null)
            {
                objConn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\curso_de_programacao\Preferencias.accdb");
            }
            return objConn;
        }

        public static void AbreConexao()
        {
            if (getConexao().State == System.Data.ConnectionState.Closed)
            {
                objConn.Open();
            }
        }
        public static void FechaConexao()
        {
            if (getConexao().State == System.Data.ConnectionState.Open)
            {
                objConn.Close();

                objConn.Dispose();

                objConn = null;
            }
        }
    }
}
