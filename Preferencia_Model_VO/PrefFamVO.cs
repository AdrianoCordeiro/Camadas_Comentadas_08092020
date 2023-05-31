using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Preferencia_Model_VO
{
    public class PrefFamVO
    {
        private FamiliaresVO ObjfamiliarVO;
        private PreferenciaVO ObjPreferenciaVO;
        private float intensidade;
        private string observacao;

        public PrefFamVO()
        {

        }
        public PrefFamVO(FamiliaresVO objParFamiliarVO, PreferenciaVO objParPreferenciaVO, float fltIntensidade, string strObservacao = null)
        {
            setObjFamiliarVO(objParFamiliarVO);
            setObjPreferenciaVO(objParPreferenciaVO);
            setIntensidade(fltIntensidade);
            setObservacao(strObservacao);
        }

        public FamiliaresVO getObjFamiliarVO()
        {
            return this.ObjfamiliarVO;
        }
        public PreferenciaVO getObjPreferenciaVO()
        {
            return this.ObjPreferenciaVO;
        }
        public float getIntensiadade()
        {
            return this.intensidade;
        }
        public string getObservacao()
        {
            return this.observacao;
        }

        public void setObjFamiliarVO(FamiliaresVO objFamiliar)
        {
        this.ObjfamiliarVO = objFamiliar;
        }
        public void setObjPreferenciaVO(PreferenciaVO objPreferencia)
        {
            this.ObjPreferenciaVO = objPreferencia;
        }
        public void setIntensidade(float fltIntensidade)
        {
            this.intensidade = fltIntensidade;
        }
        public void setObservacao(string strObservacao)
        {
            this.observacao = strObservacao;
        }

        public FamiliaresVO FamiliarVO
        {
            get {return this.ObjfamiliarVO ;}
            set {this.ObjfamiliarVO = value;}
        }
        public PreferenciaVO PreferenciaVO
        {
            get {return this.ObjPreferenciaVO ;}
            set {this.ObjPreferenciaVO = value;}
        }
        public float Intensidade
        {
            get {return this.intensidade ;}
            set {this.intensidade = value;}
        }
        public string Observacao
        {
            get {return this.observacao ;}
            set {this.observacao = value;}
        }

        public List<PrefFamVO> PrefFamCollection = new List<PrefFamVO>();
    }
}
