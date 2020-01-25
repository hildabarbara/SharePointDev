using System;
using System.Web;
using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;

namespace BaseAssinantesWeb
{
    /// <summary>
    /// Encapsulates all the information from SharePoint.
    /// </summary>
    public abstract class SharePointContext
    {
        public const string SPHostUrlKey = "SPHostUrl";
        public const string SPAppWebUrlKey = "SPAppWebUrl";
        public const string SPLanguageKey = "SPLanguage";
        public const string SPClientTagKey = "SPClientTag";
        public const string SPProductNumberKey = "SPProductNumber";

        protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);

        private readonly Uri spHostUrl;
        private readonly Uri spAppWebUrl;
        private readonly string spLanguage;
        private readonly string spClientTag;
        private readonly string spProductNumber;

        // <AccessTokenString, UtcExpiresOn>
        protected Tuple<string, DateTime> userAccessTokenForSPHost;
        protected Tuple<string, DateTime> userAccessTokenForSPAppWeb;
        protected Tuple<string, DateTime> appOnlyAccessTokenForSPHost;
        protected Tuple<string, DateTime> appOnlyAccessTokenForSPAppWeb;

        /// <summary>
        /// Gets the SharePoint host url from QueryString of the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The specified HTTP request.</param>
        /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        public static Uri GetSPHostUrl(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
            {
                throw new ArgumentNullException("httpRequest");
            }

            string spHostUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPHostUrlKey]);
            Uri spHostUrl;
            if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
            {
                return spHostUrl;
            }

            return null;
        }

        /// <summary>
        /// Gets the SharePoint host url from QueryString of the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The specified HTTP request.</param>
        /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        public static Uri GetSPHostUrl(HttpRequest httpRequest)
        {
            return GetSPHostUrl(new HttpRequestWrapper(httpRequest));
        }

        /// <summary>
        /// The SharePoint host url.
        /// </summary>
        public Uri SPHostUrl
        {
            get { return this.spHostUrl; }
        }

        /// <summary>
        /// The SharePoint app web url.
        /// </summary>
        public Uri SPAppWebUrl
        {
            get { return this.spAppWebUrl; }
        }

        /// <summary>
        /// The SharePoint language.
        /// </summary>
        public string SPLanguage
        {
            get { return this.spLanguage; }
        }

        /// <summary>
        /// The SharePoint client tag.
        /// </summary>
        public string SPClientTag
        {
            get { return this.spClientTag; }
        }

        /// <summary>
        /// The SharePoint product number.
        /// </summary>
        public string SPProductNumber
        {
            get { return this.spProductNumber; }
        }

        /// <summary>
        /// The user access token for the SharePoint host.
        /// </summary>
        public abstract string UserAccessTokenForSPHost
        {
            get;
        }

        /// <summary>
        /// The user access token for the SharePoint app web.
        /// </summary>
        public abstract string UserAccessTokenForSPAppWeb
        {
            get;
        }

        /// <summary>
        /// The app only access token for the SharePoint host.
        /// </summary>
        public abstract string AppOnlyAccessTokenForSPHost
        {
            get;
        }

        /// <summary>
        /// The app only access token for the SharePoint app web.
        /// </summary>
        public abstract string AppOnlyAccessTokenForSPAppWeb
        {
            get;
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="spHostUrl">The SharePoint host url.</param>
        /// <param name="spAppWebUrl">The SharePoint app web url.</param>
        /// <param name="spLanguage">The SharePoint language.</param>
        /// <param name="spClientTag">The SharePoint client tag.</param>
        /// <param name="spProductNumber">The SharePoint product number.</param>
        protected SharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber)
        {
            if (spHostUrl == null)
            {
                throw new ArgumentNullException("spHostUrl");
            }

            if (string.IsNullOrEmpty(spLanguage))
            {
                throw new ArgumentNullException("spLanguage");
            }

            if (string.IsNullOrEmpty(spClientTag))
            {
                throw new ArgumentNullException("spClientTag");
            }

            if (string.IsNullOrEmpty(spProductNumber))
            {
                throw new ArgumentNullException("spProductNumber");
            }

            this.spHostUrl = spHostUrl;
            this.spAppWebUrl = spAppWebUrl;
            this.spLanguage = spLanguage;
            this.spClientTag = spClientTag;
            this.spProductNumber = spProductNumber;
        }

        /// <summary>
        /// Creates a user ClientContext for the SharePoint host.
        /// </summary>
        /// <returns>A ClientContext instance.</returns>
        public ClientContext CreateUserClientContextForSPHost()
        {
            return CreateClientContext(this.SPHostUrl, this.UserAccessTokenForSPHost);
        }

        /// <summary>
        /// Creates a user ClientContext for the SharePoint app web.
        /// </summary>
        /// <returns>A ClientContext instance.</returns>
        public ClientContext CreateUserClientContextForSPAppWeb()
        {
            return CreateClientContext(this.SPAppWebUrl, this.UserAccessTokenForSPAppWeb);
        }

        /// <summary>
        /// Creates app only ClientContext for the SharePoint host.
        /// </summary>
        /// <returns>A ClientContext instance.</returns>
        public ClientContext CreateAppOnlyClientContextForSPHost()
        {
            return CreateClientContext(this.SPHostUrl, this.AppOnlyAccessTokenForSPHost);
        }

        /// <summary>
        /// Creates an app only ClientContext for the SharePoint app web.
        /// </summary>
        /// <returns>A ClientContext instance.</returns>
        public ClientContext CreateAppOnlyClientContextForSPAppWeb()
        {
            return CreateClientContext(this.SPAppWebUrl, this.AppOnlyAccessTokenForSPAppWeb);
        }

        /// <summary>
        /// Gets the database connection string from SharePoint for autohosted app.
        /// This method is deprecated because the autohosted option is no longer available.
        /// </summary>
        [ObsoleteAttribute("This method is deprecated because the autohosted option is no longer available.", true)]
        public string GetDatabaseConnectionString()
        {
            throw new NotSupportedException("This method is deprecated because the autohosted option is no longer available.");
        }

        /// <summary>
        /// Determines if the specified access token is valid.
        /// It considers an access token as not valid if it is null, or it has expired.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <returns>True if the access token is valid.</returns>
        protected static bool IsAccessTokenValid(Tuple<string, DateTime> accessToken)
        {
            return accessToken != null &&
                   !string.IsNullOrEmpty(accessToken.Item1) &&
                   accessToken.Item2 > DateTime.UtcNow;
        }

        /// <summary>
        /// Creates a ClientContext with the specified SharePoint site url and the access token.
        /// </summary>
        /// <param name="spSiteUrl">The site url.</param>
        /// <param name="accessToken">The access token.</param>
        /// <returns>A ClientContext instance.</returns>
        private static ClientContext CreateClientContext(Uri spSiteUrl, string accessToken)
        {
            if (spSiteUrl != null && !string.IsNullOrEmpty(accessToken))
            {
                return TokenHelper.GetClientContextWithAccessToken(spSiteUrl.AbsoluteUri, accessToken);
            }

            return null;
        }
    }

    /// <summary>
    /// Provides SharePointContext instances.
    /// </summary>
    public abstract class SharePointContextProvider
    {
        private static SharePointContextProvider current;

        /// <summary>
        /// The current SharePointContextProvider instance.
        /// </summary>
        public static SharePointContextProvider Current
        {
            get { return SharePointContextProvider.current; }
        }

        /// <summary>
        /// Initializes the default SharePointContextProvider instance.
        /// </summary>
        static SharePointContextProvider()
        {
            if (!TokenHelper.IsHighTrustApp())
            {
                SharePointContextProvider.current = new SharePointAcsContextProvider();
            }
            else
            {
                SharePointContextProvider.current = new SharePointHighTrustContextProvider();
            }
        }

        /// <summary>
        /// Registers the specified SharePointContextProvider instance as current.
        /// It should be called by Application_Start() in Global.asax.
        /// </summary>
        /// <param name="provider">The SharePointContextProvider to be set as current.</param>
        public static void Register(SharePointContextProvider provider)
        {
            if (provider == null)
            {
                throw new ArgumentNullException("provider");
            }

            SharePointContextProvider.current = provider;
        }

        /// <summary>
        /// Checks if it is necessary to redirect to SharePoint for user to authenticate.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
        /// <returns>Redirection status.</returns>
        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException("httpContext");
            }

            redirectUrl = null;
            bool contextTokenExpired = false;

            try
            {
                if (SharePointContextProvider.Current.GetSharePointContext(httpContext) != null)
                {
                    return RedirectionStatus.Ok;
                }
            }
            catch (SecurityTokenExpiredException)
            {
                contextTokenExpired = true;
            }

            const string SPHasRedirectedToSharePointKey = "SPHasRedirectedToSharePoint";

            if (!string.IsNullOrEmpty(httpContext.Request.QueryString[SPHasRedirectedToSharePointKey]) && !contextTokenExpired)
            {
                return RedirectionStatus.CanNotRedirect;
            }

            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);

            if (spHostUrl == null)
            {
                return RedirectionStatus.CanNotRedirect;
            }

            if (StringComparer.OrdinalIgnoreCase.Equals(httpContext.Request.HttpMethod, "POST"))
            {
                return RedirectionStatus.CanNotRedirect;
            }

            Uri requestUrl = httpContext.Request.Url;

            var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);

            // Removes the values that are included in {StandardTokens}, as {StandardTokens} will be inserted at the beginning of the query string.
            queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
            queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
            queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);

            // Adds SPHasRedirectedToSharePoint=1.
            queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1");

            UriBuilder returnUrlBuilder = new UriBuilder(requestUrl);
            returnUrlBuilder.Query = queryNameValueCollection.ToString();

            // Inserts StandardTokens.
            const string StandardTokens = "{StandardTokens}";
            string returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
            returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, StandardTokens + "&");

            // Constructs redirect url.
            string redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.AbsoluteUri, Uri.EscapeDataString(returnUrlString));

            redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);

            return RedirectionStatus.ShouldRedirect;
        }

        /// <summary>
        /// Checks if it is necessary to redirect to SharePoint for user to authenticate.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
        /// <returns>Redirection status.</returns>
        public static RedirectionStatus CheckRedirectionStatus(HttpContext httpContext, out Uri redirectUrl)
        {
            return CheckRedirectionStatus(new HttpContextWrapper(httpContext), out redirectUrl);
        }

        /// <summary>
        /// Creates a SharePointContext instance with the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The HTTP request.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
        public SharePointContext CreateSharePointContext(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
            {
                throw new ArgumentNullException("httpRequest");
            }

            // SPHostUrl
            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpRequest);
            if (spHostUrl == null)
            {
                return null;
            }

            // SPAppWebUrl
            string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SharePointContext.SPAppWebUrlKey]);
            Uri spAppWebUrl;
            if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) ||
                !(spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps))
            {
                spAppWebUrl = null;
            }

            // SPLanguage
            string spLanguage = httpRequest.QueryString[SharePointContext.SPLanguageKey];
            if (string.IsNullOrEmpty(spLanguage))
            {
                return null;
            }

            // SPClientTag
            string spClientTag = httpRequest.QueryString[SharePointContext.SPClientTagKey];
            if (string.IsNullOrEmpty(spClientTag))
            {
                return null;
            }

            // SPProductNumber
            string spProductNumber = httpRequest.QueryString[SharePointContext.SPProductNumberKey];
            if (string.IsNullOrEmpty(spProductNumber))
            {
                return null;
            }

            return CreateSharePointContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, httpRequest);
        }

        /// <summary>
        /// Creates a SharePointContext instance with the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The HTTP request.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
        public SharePointContext CreateSharePointContext(HttpRequest httpRequest)
        {
            return CreateSharePointContext(new HttpRequestWrapper(httpRequest));
        }

        /// <summary>
        /// Gets a SharePointContext instance associated with the specified HTTP context.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if not found and a new instance can't be created.</returns>
        public SharePointContext GetSharePointContext(HttpContextBase httpContext)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException("httpContext");
            }

            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
            if (spHostUrl == null)
            {
                return null;
            }

            SharePointContext spContext = LoadSharePointContext(httpContext);

            if (spContext == null || !ValidateSharePointContext(spContext, httpContext))
            {
                spContext = CreateSharePointContext(httpContext.Request);

                if (spContext != null)
                {
                    SaveSharePointContext(spContext, httpContext);
                }
            }

            return spContext;
        }

        /// <summary>
        /// Gets a SharePointContext instance associated with the specified HTTP context.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if not found and a new instance can't be created.</returns>
        public SharePointContext GetSharePointContext(HttpContext httpContext)
        {
            return GetSharePointContext(new HttpContextWrapper(httpContext));
        }

        /// <summary>
        /// Creates a SharePointContext instance.
        /// </summary>
        /// <param name="spHostUrl">The SharePoint host url.</param>
        /// <param name="spAppWebUrl">The SharePoint app web url.</param>
        /// <param name="spLanguage">The SharePoint language.</param>
        /// <param name="spClientTag">The SharePoint client tag.</param>
        /// <param name="spProductNumber">The SharePoint product number.</param>
        /// <param name="httpRequest">The HTTP request.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
        protected abstract SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest);

        /// <summary>
        /// Validates if the given SharePointContext can be used with the specified HTTP context.
        /// </summary>
        /// <param name="spContext">The SharePointContext.</param>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>True if the given SharePointContext can be used with the specified HTTP context.</returns>
        protected abstract bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext);

        /// <summary>
        /// Loads the SharePointContext instance associated with the specified HTTP context.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if not found.</returns>
        protected abstract SharePointContext LoadSharePointContext(HttpContextBase httpContext);

        /// <summary>
        /// Saves the specified SharePointContext instance associated with the specified HTTP context.
        /// <c>null</c> is accepted for clearing the SharePointContext instance associated with the HTTP context.
        /// </summary>
        /// <param name="spContext">The SharePointContext instance to be saved, or <c>null</c>.</param>
        /// <param name="httpContext">The HTTP context.</param>
        protected abstract void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext);
    }
}