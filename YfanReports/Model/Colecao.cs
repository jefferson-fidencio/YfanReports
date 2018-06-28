using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VotacaoEstampas.Model
{
    public class Colecao
    {
        public string Nome;
        public DateTime Data;
        public List<Estampa> Estampas;
        public List<Votacao> Votacoes;
    }
}
