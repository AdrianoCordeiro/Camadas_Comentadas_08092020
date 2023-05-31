using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace PreferencciaDAO_DAL
{
    // impoe um minino de Metedos a serem criados
    public abstract class DAO_DAL : DB_DAO
    {
        public abstract DataTable ConsultarBd(Object objvo_VO);

        public abstract void ConsultarBd(ref Object objvo_VO);

        public abstract bool IncluirBd(Object objvo_VO);

        public abstract bool ExcluirBd(Object objvo_VO);

        public abstract bool AlterarBd(Object objvo_VO);
    }
}
