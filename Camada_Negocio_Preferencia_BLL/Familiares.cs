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
    public class Familiares
    {
        FamiliaresFD objFamiliarFD;

        public DataTable ConsultarBd(FamiliaresVO objparFamiliarVO)
        {
            try
            {
                objFamiliarFD = new FamiliaresFD();
                return objFamiliarFD.ConsultarBd(objparFamiliarVO);
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
                objFamiliarFD = new FamiliaresFD();
                objFamiliarFD.ConsultarBd(ref objparFamiliarVO);
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
                objFamiliarFD = new FamiliaresFD();
                return objFamiliarFD.IncluirBd(objparFamiliarVO);
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
                objFamiliarFD = new FamiliaresFD();
                return objFamiliarFD.ExcluirBd(objparFamiliarVO);
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
                objFamiliarFD = new FamiliaresFD();
                return objFamiliarFD.AlterarBd(objparFamiliarVO);
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
                objFamiliarFD = new FamiliaresFD();

                objFamiliarFD.GeraExcelDoAccessPorinterop(strnNomePlanilha);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
