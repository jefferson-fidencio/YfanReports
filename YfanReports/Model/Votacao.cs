using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VotacaoEstampas.Model
{
    public class Votacao
    {
        public Cliente Cliente;
        public DateTime Data;
        public List<bool> Votos;
    }
}
