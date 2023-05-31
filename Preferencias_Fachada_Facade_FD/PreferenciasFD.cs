using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PreferencciaDAO_DAL;
using System.Data;
using Preferencia_Model_VO;


//Fachada eh a camada responsavel por traduzir novos metodos de sistemas legados ou antigo e trazer 
// para a DAO sem a necessidade de recontrucao do sistema
// Fachasda serve pra diminuir a complexibilidade do povoamento de objetos REFLEXION

namespace Preferencias_Fachada_Facade_FD
{
    //camada de fachada - facade de protecao da preferfencias DAO
    public class PreferenciasFD
    {
        PreferenciaDAO objPreferenciasDAO;


        public List<string> ImportarBdConectado()
        {
            try
            {
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.ImportarBdConectado();
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
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.ImportarBdDesconectado();
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
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.ConsultarBd(objparPreferenciaVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        } 
   
        public void ConsultarBd(ref PreferenciaVO objparPreferenciaVO)
        {
            try
            {
                objPreferenciasDAO = new PreferenciaDAO();

                Object objvo_VO = (PreferenciaVO)objparPreferenciaVO;

                objPreferenciasDAO.ConsultarBd(ref objvo_VO);
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
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.IncluirBd(objparPreferenciaVO);
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
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.ExcluirBd(objparPreferenciaVO);
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
                objPreferenciasDAO = new PreferenciaDAO();
                return objPreferenciasDAO.AlterarBd(objparPreferenciaVO);
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
                objPreferenciasDAO = new PreferenciaDAO();

                objPreferenciasDAO.GeraExcelDoAccessPorinterop(strnNomePlanilha);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
