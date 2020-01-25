using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Configuration;
//using Portal.IntegradorX;
using BaseAssinantesWeb.Models;

namespace BaseAssinantesWeb.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private string _urlRaiz = ConfigurationManager.AppSettings["IntX:UrlRaiz"];
        private string _opCode = ConfigurationManager.AppSettings["IntX:OpCode"];
        private string _systemId = ConfigurationManager.AppSettings["IntX:SystemId"];
        private string _dominio = ConfigurationManager.AppSettings["IntX:AplicacaoCliente"];

        // GET: Home
        public ActionResult Index()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    Security security = new Security();
                    var identity = security.GetIdentityClaims(User.Identity);

                    var urlRaiz = new Uri(_urlRaiz);
                    var opCode = Convert.ToInt32(_opCode);
                    var systemId = Convert.ToInt32(_systemId);

                    var integrador = new IntegradorX(urlRaiz, opCode, systemId, _dominio);

                    string urlRelativaImportarPlanilha = "faces/pages/Importacao.html&sso=1";

                    var dadosIntegracaoImportarPlanilha = integrador.ObterDadosIntegracao(identity, urlRelativaImportarPlanilha);

                    BaseAssinantesModel model = new BaseAssinantesModel()
                    {
                        EnderecoImportarPlanilha = dadosIntegracaoImportarPlanilha.Endereco,
                        ParametrosSSO = dadosIntegracaoImportarPlanilha.ParametrosSSO,
                        Token = dadosIntegracaoImportarPlanilha.Token
                    };


                }
            }

            return View();
        }
    }
}