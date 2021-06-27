using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using TCore.MsalWeb;

namespace ArbWeb
{
	public class Auth : WebApiInterop.IAccessTokenProvider
	{
		public struct AuthInfo
		{
			public string Identity;
			public string Tenant;
		};

		private string m_sClientID = "9a58a71e-183d-4a6e-9b12-251e56b70901";
        private string[] m_rgsScopes;
		private IPublicClientApplication m_app;
		private AuthInfo m_authInfo;

		private WebApiInterop.IAccessTokenProvider m_accessTokenProviderImplementation;

        private string m_sAccessToken;

        public Auth(string sClientID, string[] rgsScopes)
        {
            m_sClientID = sClientID;
            m_rgsScopes = rgsScopes;
            m_authInfo = new AuthInfo() { Identity = null, Tenant = null };
            
            m_app = PublicClientApplicationBuilder.Create(m_sClientID)
	            .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
	            .WithAuthority(AzureCloudInstance.AzurePublic, "common")
	            .Build();
        }

        public bool IsLoggedIn => m_authInfo.Identity != null;

        public void Logout()
        {
            m_authInfo = new AuthInfo();
        }

        public AuthInfo UserInfo => m_authInfo;

        public async Task<bool> Login()
        {
            AuthenticationResult result;
            IEnumerable<IAccount> accounts = await m_app.GetAccountsAsync();

            if (accounts == null)
                throw new Exception("no accounts!");

            try
            {
                result = await m_app.AcquireTokenSilent(m_rgsScopes, accounts.FirstOrDefault()).ExecuteAsync();
                //result = await m_app.AcquireTokenSilentAsync(m_rgsScopes, accounts.FirstOrDefault());
            }
            catch (MsalUiRequiredException ex)
            {
                result = await m_app.AcquireTokenInteractive(m_rgsScopes).ExecuteAsync();
                //result = await m_app.AcquireTokenAsync(m_rgsScopes);
            }

            m_authInfo = new AuthInfo() { Identity = result.Account.Username, Tenant = result.TenantId };

            // we've populated the token cache now...
            return true;
        }

        public async Task<string> GetAccessTokenAsync()
        {
            return await GetAccessTokenForScopesAsync(m_rgsScopes);
        }

        public async Task<string> GetAccessTokenForScopesAsync(string[] rgsScopes)
        {
            AuthenticationResult result;
            IEnumerable<IAccount> accounts = await m_app.GetAccountsAsync();

            if (accounts == null)
                throw new Exception("no accounts!");

            result = await m_app.AcquireTokenSilent(rgsScopes, accounts.FirstOrDefault()).ExecuteAsync();
            // result = await m_app.AcquireTokenSilentAsync(rgsScopes, accounts.FirstOrDefault());

            return result.AccessToken;
        }

        public string GetAccessToken()
        {
            Task<string> tskToken = GetAccessTokenAsync();
            tskToken.Wait();

            return tskToken.Result;
        }

        public string GetAccessTokenForScope(string[] rgsScopes)
        {
            Task<string> tskToken = GetAccessTokenForScopesAsync(rgsScopes);
            tskToken.Wait();

            return tskToken.Result;
        }
    }
}
