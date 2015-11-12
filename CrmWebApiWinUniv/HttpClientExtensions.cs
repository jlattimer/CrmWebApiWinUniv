using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace CrmWebApiWinUniv
{
	public static class HttpClientExtensions
	{
		public static Task<HttpResponseMessage> SendAsJsonAsync<T>(this HttpClient client, HttpMethod method, string requestUri, T value)
		{
			var content = value.GetType().Name.Equals("JObject") ? 
				value.ToString() : 
				JsonConvert.SerializeObject(value, new JsonSerializerSettings() { DefaultValueHandling = DefaultValueHandling.Ignore });

			HttpRequestMessage request = new HttpRequestMessage(method, requestUri) { Content = new StringContent(content) };
			request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

			return client.SendAsync(request);
		}
	}
}
