using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Preferencia_Model_VO;
using PreferencciaDAO_DAL;

namespace Preferencias_Fachada_Facade_FD
{
    public class FamiliaresFD
    {
        FamiliaresDAO objFamiliarDAO;
        

        public DataTable ConsultarBd(FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarDAO = new FamiliaresDAO();

                return objFamiliarDAO.ConsultarBd(objparFamiliarVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void ConsultarBd(ref FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarDAO = new FamiliaresDAO();

                Object objvo_VO = (FamiliaresVO)objparFamiliarVO;

                objFamiliarDAO.ConsultarBd( ref objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBd(FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarDAO = new FamiliaresDAO();

                return objFamiliarDAO.IncluirBd(objparFamiliarVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBd(FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarDAO = new FamiliaresDAO();

                return objFamiliarDAO.ExcluirBd(objparFamiliarVO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBd(FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarDAO = new FamiliaresDAO();

                return objFamiliarDAO.AlterarBd(objparFamiliarVO);
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
                objFamiliarDAO = new FamiliaresDAO();

                objFamiliarDAO.GeraExcelDoAccessPorinterop(strnNomePlanilha);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
