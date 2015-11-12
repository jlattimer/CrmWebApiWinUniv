using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Windows.Security.Authentication.Web;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.UI.Popups;
using Windows.UI.Xaml;

namespace CrmWebApiWinUniv
{
	/// <summary>
	/// An empty page that can be used on its own or navigated to within a Frame.
	/// </summary>
	public sealed partial class MainPage
	{
		//This was registered in Azure AD as a WEB APPLICATION AND/OR WEB API

		//Azure Application Client ID
		private const string _clientId = "00000000-0000-0000-0000-000000000000";
		// Azure Application REPLY URL - generate from included GetAppRedirectURI function
		//private const string _redirectUrl = "MS-APPX-WEB://MICROSOFT.AAD.BROKERPLUGIN/S-0-00-0-000000000-000000000-0000000000-000000000-0000000000-0000000000-000000000";
		//CRM URL
		private const string _resource = "https://org.crm.dynamics.com";
		//Azure Directory OAUTH 2.0 AUTHORIZATION ENDPOINT
		private const string _authority = "https://login.microsoftonline.com/00000000-0000-0000-0000-000000000000";

		private static string _accessToken;

		public MainPage()
		{
			InitializeComponent();
		}

		private async void MainPage_OnLoaded(object sender, RoutedEventArgs e)
		{
			//var redirect = GetAppRedirectURI();

			try
			{
				WebAccountProvider wap =
						await WebAuthenticationCoreManager.FindAccountProviderAsync("https://login.microsoft.com", _authority);

				WebTokenRequest wtr = new WebTokenRequest(wap, string.Empty, _clientId);
				wtr.Properties.Add("resource", _resource);
				WebTokenRequestResult wtrr = await WebAuthenticationCoreManager.RequestTokenAsync(wtr);

				if (wtrr.ResponseStatus == WebTokenRequestStatus.Success)
					_accessToken = wtrr.ResponseData[0].Token;
			}
			catch (Exception ex)
			{
				ShowException(ex);
			}
		}

		private async void InputButton_OnClick(object sender, RoutedEventArgs e)
		{
			try
			{
				string userId = await Task.Run(WhoAmI);
				if (!string.IsNullOrEmpty(userId))
					whoAmIOutput.Text = userId;

				if (!string.IsNullOrEmpty(userId))
				{
					string fullName = await Task.Run(() => Retrieve(userId));
					retrieveOutput.Text = "Fullname: " + fullName;
				}

				string accountId = await Task.Run(() => Create("Windows Universal Test"));
				if (!string.IsNullOrEmpty(accountId))
					createOutput.Text = "Created: " + accountId;

				accountId = await Task.Run(() => Update(accountId));
				if (!string.IsNullOrEmpty(accountId))
					updateOutput.Text = "Updated: " + accountId;

				accountId = await Task.Run(() => Delete(accountId));
				if (!string.IsNullOrEmpty(accountId))
					deleteOutput.Text = "Deleted: " + accountId;
			}
			catch (Exception ex)
			{
				ShowException(ex);
			}
		}

		public async Task<string> WhoAmI()
		{
			try
			{
				using (HttpClient httpClient = new HttpClient())
				{
					httpClient.BaseAddress = new Uri(_resource);
					httpClient.Timeout = new TimeSpan(0, 2, 0);
					httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
					httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
					httpClient.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));
					httpClient.DefaultRequestHeaders.Authorization =
						new AuthenticationHeaderValue("Bearer", _accessToken);

					//Unbound Function
					//The URL will change in 2016 to include the API version - api/data/v8.0/WhoAmI
					HttpResponseMessage whoAmIResponse =
						await httpClient.GetAsync("api/data/WhoAmI");

					if (!whoAmIResponse.IsSuccessStatusCode)
						return null;

					JObject jWhoAmIResponse =
						JObject.Parse(whoAmIResponse.Content.ReadAsStringAsync().Result);
					return jWhoAmIResponse["UserId"].ToString();
				}
			}
			catch (Exception ex)
			{
				ShowException(ex);
				return null;
			}
		}

		public async Task<string> Retrieve(string userId)
		{
			try
			{
				using (HttpClient httpClient = new HttpClient())
				{
					httpClient.BaseAddress = new Uri(_resource);
					httpClient.Timeout = new TimeSpan(0, 2, 0);
					httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
					httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
					httpClient.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));
					httpClient.DefaultRequestHeaders.Authorization =
						new AuthenticationHeaderValue("Bearer", _accessToken);

					//Retrieve 
					//The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
					HttpResponseMessage retrieveResponse =
						await httpClient.GetAsync("api/data/systemusers(" +
						userId + ")?$select=fullname");

					if (!retrieveResponse.IsSuccessStatusCode)
						return null;

					JObject jRetrieveResponse =
						JObject.Parse(retrieveResponse.Content.ReadAsStringAsync().Result);
					return jRetrieveResponse["fullname"].ToString();
				}
			}
			catch (Exception ex)
			{
				ShowException(ex);
				return null;
			}
		}

		public async Task<string> Create(string name)
		{
			try
			{
				using (HttpClient httpClient = new HttpClient())
				{
					httpClient.BaseAddress = new Uri(_resource);
					httpClient.Timeout = new TimeSpan(0, 2, 0);
					httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
					httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
					httpClient.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));
					httpClient.DefaultRequestHeaders.Authorization =
						new AuthenticationHeaderValue("Bearer", _accessToken);

					//Create
					JObject newAccount = new JObject
					{
						{"name", name},
						{"telephone1", "111-888-7777"}
					};

					//The URL will change in 2016 to include the API version - api/data/v8.0/accounts
					HttpResponseMessage createResponse =
						await httpClient.SendAsJsonAsync(HttpMethod.Post, "api/data/accounts", newAccount);

					if (!createResponse.IsSuccessStatusCode)
						return null;

					string accountUri = createResponse.Headers.GetValues("OData-EntityId").FirstOrDefault();
					return accountUri != null ?
						Guid.Parse(accountUri.Split('(', ')')[1]).ToString() :
						null;
				}
			}
			catch (Exception ex)
			{
				ShowException(ex);
				return null;
			}
		}

		public async Task<string> Update(string accountId)
		{
			try
			{
				using (HttpClient httpClient = new HttpClient())
				{
					httpClient.BaseAddress = new Uri(_resource);
					httpClient.Timeout = new TimeSpan(0, 2, 0);
					httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
					httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
					httpClient.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));
					httpClient.DefaultRequestHeaders.Authorization =
						new AuthenticationHeaderValue("Bearer", _accessToken);

					//Update 
					JObject account = new JObject
					{
						{"websiteurl", "http://www.microsoft.com"}
					};

					//The URL will change in 2016 to include the API version - api/data/v8.0/accounts
					HttpResponseMessage updateResponse =
						await httpClient.SendAsJsonAsync(new HttpMethod("PATCH"), "api/data/accounts(" + accountId + ")", account);
					return updateResponse.IsSuccessStatusCode ? accountId : null;
				}
			}
			catch (Exception ex)
			{
				ShowException(ex);
				return null;
			}
		}

		public async Task<string> Delete(string accountId)
		{
			try
			{
				using (HttpClient httpClient = new HttpClient())
				{
					httpClient.BaseAddress = new Uri(_resource);
					httpClient.Timeout = new TimeSpan(0, 2, 0);
					httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
					httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
					httpClient.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));
					httpClient.DefaultRequestHeaders.Authorization =
						new AuthenticationHeaderValue("Bearer", _accessToken);

					//Delete
					//The URL will change in 2016 to include the API version - api/data/v8.0/accounts
					HttpResponseMessage deleteResponse =
						await httpClient.DeleteAsync("api/data/accounts(" + accountId + ")");

					return deleteResponse.IsSuccessStatusCode ? accountId : null;
				}
			}
			catch (Exception ex)
			{
				ShowException(ex);
				return null;
			}
		}

		private async void ShowException(Exception ex)
		{
			var dialog = new MessageDialog(ex.Message + Environment.NewLine + ex.StackTrace);
			await dialog.ShowAsync();
		}

		private static string GetAppRedirectURI()
		{
			// Windows 10 universal apps require redirect URI in the format below. Add a breakpoint to this line and run the app before you register it, so that
			// you can supply the correct redirect URI value.
			return string.Format("ms-appx-web://microsoft.aad.brokerplugin/{0}", WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host).ToUpper();
		}
	}
}
