using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using PreferencciaDAO_DAL;
using Preferencia_Model_VO;

namespace Preferencias_Fachada_Facade_FD
{
    public class PrefFamFD
    {
        PrefFamDAO objPrefFamDAO;
        

        public DataTable ConsultarBd(PrefFamVO objVo_VO)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();
                return objPrefFamDAO.ConsultarBd(objVo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ConsultarBd(ref PrefFamVO objVo_VO)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();
                Object objparPrefFamVO = (Object)objVo_VO; 
                objPrefFamDAO.ConsultarBd(ref objparPrefFamVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBd(PrefFamVO objvo_VO)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();
                return objPrefFamDAO.IncluirBd(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBd(PrefFamVO objvo_VO)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();
                return objPrefFamDAO.ExcluirBd(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBd(PrefFamVO objvo_VO)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();
                return objPrefFamDAO.AlterarBd(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void GeraExcelDoAccessPorinterop(string strnNomePlanilha, int intCod)
        {
            try
            {
                objPrefFamDAO = new PrefFamDAO();

                objPrefFamDAO.GeraExcelDoAccessPorinterop(strnNomePlanilha, intCod);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
