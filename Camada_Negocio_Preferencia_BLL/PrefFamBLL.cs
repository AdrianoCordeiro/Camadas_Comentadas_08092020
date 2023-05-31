using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Preferencia_Model_VO;
using Preferencias_Fachada_Facade_FD;

namespace Camada_Negocio_Preferencia_BLL
{
    public class PrefFamBLL
    {
        PrefFamFD objPrefFamFD;

        public DataTable ConsultarBd(PrefFamVO objVo_VO)
        {
            try
            {
                objPrefFamFD = new PrefFamFD();
                return objPrefFamFD.ConsultarBd(objVo_VO);
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
                objPrefFamFD = new PrefFamFD();
                objPrefFamFD.ConsultarBd(ref objVo_VO);
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
                objPrefFamFD = new PrefFamFD();
                return objPrefFamFD.IncluirBd(objvo_VO);
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
                objPrefFamFD = new PrefFamFD();
                return objPrefFamFD.ExcluirBd(objvo_VO);
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
                objPrefFamFD = new PrefFamFD();
                return objPrefFamFD.AlterarBd(objvo_VO);
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
                objPrefFamFD = new PrefFamFD();

                objPrefFamFD.GeraExcelDoAccessPorinterop(strnNomePlanilha, intCod);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
