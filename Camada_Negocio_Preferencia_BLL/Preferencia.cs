using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Preferencia_Model_VO;
using Preferencias_Fachada_Facade_FD;

// (BLL - Bussines Logical Layer) - Camada Logica de Negocios - Biblioteca de classes ou DLL Dynamic Linked
// Library). A camada de negocios responsavel por criar uma biblioteca de clasees que concentra as regras 
// do negocio "Preferencias" com as suas caracteristicas (atributos) e as suas operacoes (metodos) que sao 
// descritos e implementados para as funcoes principais dele. Sao as funcoes de reutilizacao, reuso e 
// portabilidade das operacoes (metodos ou funcoes) para outros sistemas e funcoes de isolamento funcional 
// para otimizar a manutencao e facilitar ao maximo as necessidades e acoes de manutencao e sustentabilidade
// do negocio "Preferencias".
namespace Camada_Negocio_Preferencia_BLL
{
    // classe de negocios preferencia que possuem as caracteristicas (atributos) 
    //e as operacoes (metodos) relativos ao manuseio e gestao de "preferencias"
    public class Preferencia
    {
       PreferenciasFD objPreferenciaFD;

        StreamReader objLeitorTxt; // cria objeto de leitura de arquivos
        string strLinhaLida; // cria variavel string para leitura do arquivo txt

        public List<string> ImportaTextoWhile() // assinatura ou radical do metodos eh a primeira linha 
        {                                       // do metodo
            // Bloco try/catch serve para tratamento de excecoes, tratamento de codigos que podem  nao ser
            // totalmente atendidos e gerarem alguma excecao/erro. 
            try // O Try consegue recuperar erros que possam ocorrer no codigo fornecido em seu bloco.
            {
                //criar variavel Para receber retorno
                List<string> resultado = new List<string>();

                objLeitorTxt=new StreamReader(@"C:\curso_de_programacao\Preferencias.txt");

                strLinhaLida = objLeitorTxt.ReadLine();

                while (strLinhaLida!=null)
                {
                   resultado.Add(strLinhaLida);
                   strLinhaLida = objLeitorTxt.ReadLine();
                }
                objLeitorTxt.Close();
                return resultado;
            }
            // O catch por sua vez faz o tratamento dos erros que acontecerem.
            catch (Exception ex)
            {
                //throw instancia e captura excessao em uma mensagem EX
                throw new Exception("Falha no Importar Texto do arquivo : " + ex.Message);
            }
        } 

        public List<string> ImportarBdConectado() 
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                return objPreferenciaFD.ImportarBdConectado();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<string> ImportarBdDesconectado()
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                return objPreferenciaFD.ImportarBdDesconectado();

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable ConsultarBd(PreferenciaVO objparPreferenciaVO)
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                return objPreferenciaFD.ConsultarBd(objparPreferenciaVO);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBd(PreferenciaVO objparPreferenciaVO)
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();


                return objPreferenciaFD.IncluirBd(objparPreferenciaVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBd(PreferenciaVO objparPreferenciaVO)
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                return objPreferenciaFD.ExcluirBd(objparPreferenciaVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBd(PreferenciaVO objparPreferenciaVO)
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                return objPreferenciaFD.AlterarBd(objparPreferenciaVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
 
        }
        public void GeraExcelDoAccessPorinterop(string strnNomePlanilha)
        {
            try
            {
                objPreferenciaFD = new PreferenciasFD();

                objPreferenciaFD.GeraExcelDoAccessPorinterop(strnNomePlanilha);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
