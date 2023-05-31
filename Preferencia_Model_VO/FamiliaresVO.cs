using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Preferencia_Model_VO
{
    public class FamiliaresVO
    {
        private int cod;
        private string nome;
        private string sexo;
        private int idade;
        private double ganhoTotalMensal;
        private double gastoTotalMensal;
        private string observacao;

        public FamiliaresVO()
        { 
        
        }

        public FamiliaresVO(int intCod, string strNome, string strSexo)
        {
            setCod(intCod);
            setNome(strNome);
            setSexo(strSexo);
        }

        public FamiliaresVO(int intCod, string strNome, string strSexo,int intIdade, double dbGanho, double dbGasto, string strObs = null )
        {
            setCod(intCod);
            setNome(strNome);
            setSexo(strSexo);
            setIdade(intIdade);
            setGanhoTotalMensal(dbGanho);
            setGastoTotalMensal(dbGasto);
            setObservacao(strObs);
        }


        //<-- GETTERS -->
        public int getCod()
        {
            return this.cod;
        }
        public string getNome()
        { 
        return this.nome;
        }
        public string getSexo()
        {
            return this.sexo;
        }
        public int getIdade()
        {
            return this.idade;
        }
        public double getGanhoTotalMensal()
        {
            return this.ganhoTotalMensal;
        }
        public double getGastoTotalMensal()
        {
            return this.gastoTotalMensal;
        }
        public string getObservacao()
        {
            return this.observacao;
        }
        //<-- SETTERS -->
        public void setCod(int intCod)
        {
            this.cod = intCod;
        }
        public void setNome(string strNome)
        {
            this.nome = strNome;
        }
        public void setSexo(string strSexo)
         {
             if (strSexo == "MASCULINO" || strSexo == "FEMININO" || strSexo == "INDEFINIDO")
             {
                 this.sexo = strSexo;
             }
             else
             {
                 throw new Exception("Atributo Sexo Inexistente!");
             }
        }
        public void setIdade(int intIdade)
        {
            this.idade = intIdade;
        }
        public void setGanhoTotalMensal(double dbGanho)
        {
            this.ganhoTotalMensal = dbGanho;
        }
        public void setGastoTotalMensal(double dbGasto)
        {
            this.gastoTotalMensal = dbGasto;
        }
        public void setObservacao(string strObs)
        {
            this.observacao = strObs;
        }

        //GETTERS E SETTERS MicroSoft
        public int COD
        { 
            get { return this.cod; }
            set { this.cod = value; }
        }
        public string Nome
        {
            get { return this.nome; }
            set { this.nome = value; }
        }
        public string Sexo
        {
            get { return this.sexo; }
            set
            {
                if (value == "MASCULINO" || value == "FEMININO" || value == "INDEFINIDO")
                {
                    this.sexo = value;
                }
                else
                {
                    throw new Exception("Atributo Sexo Inexistente!");
                }

            }
        }
        public int Idade
        {
            get { return this.idade; }
            set { this.idade = value; }
        }
        public double Ganho
        {
            get { return this.ganhoTotalMensal; }
            set { this.ganhoTotalMensal = value; }
        }
        public double Gasto
        {
            get { return this.gastoTotalMensal; }
            set { this.gastoTotalMensal = value; }
        }
        public string Obs
        {
            get { return this.observacao; }
            set { this.observacao = value; }
        }

        public List<FamiliaresVO> objFamiliaresVOCollection = new List<FamiliaresVO>();
    }
    
}
