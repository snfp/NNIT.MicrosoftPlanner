using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace UiPath.Shared.Activities.HTTP
{
    public class HTTPHandler
    {
        public async Task<string> GetRequest(string restUrl, string accessToken, CancellationToken cancellationToken)
        {
            string jsonresult = null;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var accept = "application/json";
                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    using (var response = await client.GetAsync(restUrl, cancellationToken))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            jsonresult = await response.Content.ReadAsStringAsync();
                        }
                        else
                        {
                            throw new Exception("error getting data - " + response.StatusCode.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
            return jsonresult;
        }
        public async Task<string> DeleteRequest(string restUrl, string accessToken, string etag, CancellationToken cancellationToken)
        {
            string statusCode = null;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    var accept = "application/json";
                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Add("If-Match", etag);

                    //client
                    using (var response = await client.DeleteAsync(new Uri(restUrl), cancellationToken))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            statusCode = response.StatusCode.ToString();
                        }
                        else
                        {
                            throw new Exception("Error getting data: " + response.StatusCode.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw (ex); // TODO: handle the exception
            }
            return statusCode;
        }
        public async Task<string> PostRequest(string restUrl, string accessToken, string jsonString, CancellationToken cancellationToken)
        {
            string jsonresult = null;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    var accept = "application/json";
                    client.DefaultRequestHeaders.Add("Accept", accept);

                    HttpContent httpContent = new StringContent(jsonString, Encoding.UTF8, "application/json");

                    //client
                    using (var response = await client.PostAsync(new Uri(restUrl), httpContent, cancellationToken))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            jsonresult = await response.Content.ReadAsStringAsync();
                            if (string.IsNullOrEmpty(jsonresult)) return (response.StatusCode.ToString());
                        }
                        else
                        {
                            throw new Exception("Error getting data: " + response.StatusCode.ToString());
                        }
                    }
                }
                return jsonresult;
            }
            catch (Exception ex)
            {
                // TODO: handle the exception
                Console.WriteLine(ex.Message);
                throw (ex); 
            }
            return jsonresult;
        }
        public async Task<string> PatchRequest(string restUrl, string accessToken, string etag, string json, CancellationToken cancellationToken)
        {
            string jsonresult = null;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    var accept = "application/json";
                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Add("If-Match", etag);

                    HttpContent httpContent = new StringContent(json, Encoding.UTF8, "application/json");

                    //client

                    using (var response = await client.PatchAsync(new Uri(restUrl), httpContent, cancellationToken))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            jsonresult = await response.Content.ReadAsStringAsync();
                            if (string.IsNullOrEmpty(jsonresult)) return (response.ReasonPhrase);
                        }
                        else
                        {
                            throw new Exception("Error getting data: " + response.StatusCode.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw (ex); // TODO: handle the exception
            }
            return jsonresult;
        }
    }
    public static class HttpClientEx
    {
        public static async Task<HttpResponseMessage> PatchAsync(this HttpClient client, Uri requestUri, HttpContent iContent, CancellationToken cancellationToken)
        {
            var method = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(method, requestUri)
            {
                Content = iContent
            };

            var response = default(HttpResponseMessage);
            try
            {
                response = await client.SendAsync(request, cancellationToken);
            }
            catch (TaskCanceledException e)
            {
                Console.WriteLine("ERROR: " + e.ToString());
            }

            return response;
        }
    }
}
