using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Camada MOdel (Modelo) responsavel por SIMPLIFICAR a abstracao
// de dados internos do sistema atraves da criacao de objetos
//estruturais de classes modelos que irao ser os responsaveis pela comunicacao interna entre as camadas
namespace Preferencia_Model_VO
{
    //diretivas de acesso Private, Public e Protected
    //Encapsulacao ou Encapsulamento (atributo Privado e metodos GETTER e SETTER publicos
    // atributos da classe - caracteristica dos objetos dessa classe
    public class PreferenciaVO
    {
        private int iD;
        private string descricao;// _ ou minusculo - convercao de merdado

        //metodos da classe - operacoes da classe - comportamento dos objetos da classe

        // construtores 
        // os construtores nascem da necessidade do sistema

        // construtor limpo
        public PreferenciaVO() 
        {
        
        }

        // construtor com sobrecararga (parametros)
        // em alguns casos por logica posso necessitar somente deste construtor
        // no incluir (pois eh obrigatoria passagem da descrcao)
        public PreferenciaVO(string strDescricao)
        {
            setDescricao(strDescricao);
        }

        //metodos de acesso - Getters e Setters - acessam por leitura os atributos (getters) 
        // ou acessam para alteracao do atributo(setters)

        public PreferenciaVO(int intId, string strDescricao) //utilizando os setters por seguranca por causa
        { // de possiveis verificacoes de seguranca
            setId(intId);
            setDescricao(strDescricao);
        }

        // Getters e Setters Classicos

        public int getId()
        {
            return this.iD;
        }

        
        // Getters
        public string getDescricao() 
        {
            return this.descricao;
        }

        //Setter

        public void setId(int intId)
        {
            this.iD = intId;
        }

        public void setDescricao(string strDescricao) 
        {
            // o campo Descricao configura campo privado
            // usando this para referenciar atributo
            this.descricao = strDescricao;
        }

        // Getters e Setters Microsoft - Propriedades

        // Getter & Setter
        public int ID
        {
            get { return this.iD; }
            set { this.iD = value; }
        }

        public string Descricao
        {
            get { return this.descricao; }// a direita da igualdade, assume como getter
            set { this.descricao = value; }// a esquerda da igualdade, assume como setter
        }

        // exemplo de geracao automatico de getter e setter (ms) - snniped - #propfull, #prop e similares
        //propfull
        //private int myVar;

        //public int MyProperty
        //{
        //    get { return myVar; }
        //    set { myVar = value; }
        //}
        ////prop - traz na mesma linha 
        //public int MyProperty { get; set; }

        //atributo de colecao publica para facilitar a navegacao entre os modelos e o 
        //armazenamento interno de um conjunto de modelos
        public List<PreferenciaVO> PreferenciaVOCollection = new List<PreferenciaVO>();
    }

    // outra lugar de colocar familiaresVO ou demais classes VO;
    //public class FamiliaresVO
    //{
    //    private int cod;
    //    private string nome;
    //    private string sexo;
    //    private int idade;
    //    private double ganhoTotalMensal;
    //    private double gastoTotalMensal;
    //    private string observacao;

    //    public FamiliaresVO()
    //    {

    //    }
    //    public FamiliaresVO(int intCod, string strNome, string strSexo)
    //    {
    //        setCod(intCod);
    //        setNome(strNome);
    //        setSexo(strSexo);
    //    }
    //    public FamiliaresVO(int intCod, string strNome, string strSexo, int intIdade, double dbGanho, double dbGasto, string strObs)
    //    {
    //        setCod(intCod);
    //        setNome(strNome);
    //        setSexo(strSexo);
    //        setIdade(intIdade);
    //        setGanhoTotalMensal(dbGanho);
    //        setGastoTotalMensal(dbGasto);
    //        setObservacao(strObs);
    //    }


    //    //<-- GETTERS -->
    //    public int getCod()
    //    {
    //        return this.cod;
    //    }
    //    public string getNome()
    //    {
    //        return this.nome;
    //    }
    //    public string getSexo()
    //    {
    //        return this.sexo;
    //    }
    //    public int getIdade()
    //    {
    //        return this.idade;
    //    }
    //    public double getGanhoTotalMensal()
    //    {
    //        return this.ganhoTotalMensal;
    //    }
    //    public double getGastoTotalMensal()
    //    {
    //        return this.gastoTotalMensal;
    //    }
    //    public string getObservacao()
    //    {
    //        return this.observacao;
    //    }
    //    //<-- SETTERS -->
    //    public void setCod(int intCod)
    //    {
    //        this.cod = intCod;
    //    }
    //    public void setNome(string strNome)
    //    {
    //        this.nome = strNome;
    //    }
    //    public void setSexo(string strSexo)
    //    {
    //        if (strSexo == "MASCULINO" || strSexo == "FEMININO" || strSexo == "INDEFINIDO")
    //        {
    //            this.sexo = strSexo;
    //        }
    //        else
    //        {
    //            throw new Exception("Atributo Sexo Inexistente!");
    //        }
    //    }
    //    public void setIdade(int intIdade)
    //    {
    //        this.idade = intIdade;
    //    }
    //    public void setGanhoTotalMensal(double dbGanho)
    //    {
    //        this.ganhoTotalMensal = dbGanho;
    //    }
    //    public void setGastoTotalMensal(double dbGasto)
    //    {
    //        this.gastoTotalMensal = dbGasto;
    //    }
    //    public void setObservacao(string strObs)
    //    {
    //        this.observacao = strObs;
    //    }

    //    //GETTERS E SETTERS MicroSoft
    //    public int COD
    //    {
    //        get { return this.cod; }
    //        set { this.cod = value; }
    //    }
    //    public string Nome
    //    {
    //        get { return this.nome; }
    //        set { this.nome = value; }
    //    }
    //    public string Sexo
    //    {
    //        get { return this.sexo; }
    //        set
    //        {
    //            if (value == "MASCULINO" || value == "FEMININO" || value == "INDEFINIDO")
    //            {
    //                this.sexo = value;
    //            }
    //            else
    //            {
    //                throw new Exception("Atributo Sexo Inexistente!");
    //            }

    //        }
    //    }
    //    public int Idade
    //    {
    //        get { return this.idade; }
    //        set { this.idade = value; }
    //    }
    //    public double Ganho
    //    {
    //        get { return this.ganhoTotalMensal; }
    //        set { this.ganhoTotalMensal = value; }
    //    }
    //    public double Gasto
    //    {
    //        get { return this.gastoTotalMensal; }
    //        set { this.gastoTotalMensal = value; }
    //    }
    //    public string Obs
    //    {
    //        get { return this.observacao; }
    //        set { this.observacao = value; }
    //    }

    //    public List<FamiliaresVO> objFamiliaresVOCollection = new List<FamiliaresVO>();
    //}
    
}
