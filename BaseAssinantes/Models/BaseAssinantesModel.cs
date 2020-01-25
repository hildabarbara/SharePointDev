using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BaseAssinantesWeb.Models
{
    public class BaseAssinantesModel
    {
        public Uri EnderecoImportarPlanilha { get; set; }

        public string ParametrosSSO { get; set; }

        public string Token { get; set; }
    }
}