using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Internal;
using BaseAssinantesWeb.Models;

namespace BaseAssinantesWeb
{
    public class Security
    {
        string tenantID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/tenantid").Value;
        string userObjectID = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
        string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        string appKey = ConfigurationManager.AppSettings["ida:ClientSecret"];
        string graphResourceID = "https://graph.windows.net";
        string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];

        /// <summary>
        /// Obtem o token de um usuario logado.
        /// </summary>
        private async Task<string> GetTokenForApplication()
        {
            string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            // get a token for the Graph without triggering any user interaction (from the cache, via multi-resource refresh token, etc)
            ClientCredential clientcred = new ClientCredential(clientId, appKey);
            // initialize AuthenticationContext with the token cache of the currently signed in user, as kept in the app's database
            Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext authenticationContext = new Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(aadInstance + tenantID, new ADALTokenCache(signedInUserID));
            AuthenticationResult authenticationResult = await authenticationContext.AcquireTokenSilentAsync(graphResourceID, clientcred, new UserIdentifier(userObjectID, UserIdentifierType.UniqueId));
            return authenticationResult.AccessToken;
        }

        internal object GetIdentityClaims(IIdentity identity)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Obtem a identidade do usuario com o e-mail alternaitvo da Claims do usuário.
        /// </summary>
        /// <param name="userIdentity">Identidade do usuário</param>
        public IIdentity GetIdentityClaims(IIdentity userIdentity)
        {
            var identity = userIdentity as ClaimsIdentity;
            var claimsIdentity = new ClaimsIdentity(userIdentity);
            var upnClaim = claimsIdentity.FindFirst(ClaimTypes.Upn);

            if (upnClaim == null)
                throw new ApplicationException("UserPrincipalName-upn do Claims do usuario não encontrado");

            string otherEmail = GetAlternateEmailUser(upnClaim.Value);

            identity.AddClaim(new Claim(ClaimTypes.Email, otherEmail));

            return identity;
        }

        /// <summary>
        /// Obtem o e-mail alternativo do usuario da Claims do usuário.
        /// </summary>
        /// <param name="userPrincipalName">UserprincipalName (upn da Claims do usuário)</param>
        private string GetAlternateEmailUser(string userPrincipalName)
        {
            Uri servicePointUri = new Uri(graphResourceID);
            Uri serviceRoot = new Uri(servicePointUri, tenantID);
            ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                  async () => await GetTokenForApplication());

            var result = activeDirectoryClient.Users
                    .Where(u => u.UserPrincipalName.Equals(userPrincipalName))
                    .ExecuteAsync().Result.CurrentPage.ToList();

            if (result == null)
                throw new ApplicationException("Usuario não encontrado no AD");

            IUser user = result.First();

            if (user == null)
                throw new ApplicationException("Usuario não encontrado no AD");

            if (user.OtherMails == null)
                throw new ApplicationException("E-mail do usuario não encontrado no AD");

            if (user.OtherMails.Count == 0)
                throw new ApplicationException("E-mail do usuario não encontrado no AD");

            string email = user.OtherMails.FirstOrDefault();

            if (string.IsNullOrEmpty(email))
                throw new ApplicationException("E-mail do usuario não encontrado no AD");

            return email;
        }
    }
}