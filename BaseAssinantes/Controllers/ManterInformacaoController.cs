using BaseAssinantesWeb;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Mvc;
//using Portal.IntegradorX;
//using BaseAssinantes.Models;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Utilities;
using System.Security.Claims;

namespace BaseAssinantesWeb.Controllers
{
    [Authorize]
    public class ManterInformacaoController : Controller
    {
        private string _urlRaiz = ConfigurationManager.AppSettings["Trust:UrlRaiz"];
        private string _opCode = ConfigurationManager.AppSettings["Trust:OpCode"];
        private string _systemId = ConfigurationManager.AppSettings["Trust:SystemId"];
        private string _dominio = ConfigurationManager.AppSettings["Trust:AplicacaoCliente"];

        // GET: ManterInformacao
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
                    var dominio = _dominio;

                    var integrador = new IntegradorX(urlRaiz, opCode, systemId, dominio);

                    string urlRelativaManterInformacao = "faces/pages/InformacaoOperadora.html&sso=1";

                    var dadosIntegracaoManterInformacao = integrador.ObterDadosIntegracao(identity, urlRelativaManterInformacao);

                    ManterInformacaoModel model = new ManterInformacaoModel()
                    {
                        EnderecoManterInformacao = dadosIntegracaoManterInformacao.Endereco,
                        ParametrosSSO = dadosIntegracaoManterInformacao.ParametrosSSO,
                        Token = dadosIntegracaoManterInformacao.Token
                    };

                    return View(model);
                }
            }

            return View();
        }
    }
}