using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Net;

public class PIWebApi
{
    private readonly string _urlEndpoint;
    private static readonly HttpClientHandler _clientHandler;
    private static readonly HttpClient _client;

    static PIWebApi()
    {
        _clientHandler = new HttpClientHandler
        {
            ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => true,
            Credentials = new NetworkCredential(
                Environment.GetEnvironmentVariable("PIWEBAPI_USERNAME"), 
                Environment.GetEnvironmentVariable("PIWEBAPI_PASSWORD"))
        };

        _client = new HttpClient(_clientHandler)
        {
            DefaultRequestHeaders =
            {
                Accept = { new MediaTypeWithQualityHeaderValue("application/json") },
                { "X-Requested-With", "XMLHttpRequest" }
            }
        };
    }

    public PIWebApi(string urlEndpoint)
    {
        _urlEndpoint = urlEndpoint;
    }

    public async Task<dynamic> GetPiWebApiAsync(string customUrl)
    {
        string apiUrl = $"{_urlEndpoint}/{customUrl}";
        try
        {
            HttpResponseMessage response = await _client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<dynamic>(responseBody);
        }
        catch (HttpRequestException)
        {
            return JsonConvert.DeserializeObject<dynamic>("{}");
        }
    }   

}
