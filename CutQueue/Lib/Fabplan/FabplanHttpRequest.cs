using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CutQueue.Lib.Fabplan
{
    /// <summary>
    /// A class that handles http requests to Fabplan
    /// </summary>
    public sealed class FabplanHttpRequest
    {
        private static readonly FabplanHttpRequest instance = new FabplanHttpRequest();
        private readonly HttpClient httpClient = null;
        private HttpResponseMessage response = null;

        /// <summary>
        /// Private constructor
        /// </summary>
        private FabplanHttpRequest()
        {
            httpClient = new HttpClient();
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~FabplanHttpRequest()
        {
            httpClient.Dispose();
        }

        /// <summary>
        /// Gets the static instance of the class
        /// </summary>
        /// <returns>The static instance of the class</returns>
        private static FabplanHttpRequest GetInstance()
        {
            return instance;
        }

        /// <summary>
        /// Return the HttpClient object of the class.
        /// </summary>
        /// <returns></returns>
        private HttpClient GetHttpClient()
        {
            return httpClient;
        }

        /// <summary>
        /// Performs a get http request to Fabplan
        /// </summary>
        /// <param name="url"> The url</param>
        /// <param name="parameters">The parameters used to build the query string</param>
        /// <returns>The decoded response</returns>
        public static async Task<dynamic> Get(string url, object parameters = null)
        {
            try
            {
                if (parameters != null)
                {
                    List<string> parameterStringArray = new List<string>();
                    foreach (PropertyInfo property in parameters.GetType().GetProperties())
                    {
                        parameterStringArray.Add($"{property.Name}={property.GetValue(parameters)}");
                    }
                    url += (parameterStringArray.Count() > 0) ? $"?{string.Join("&", parameterStringArray)}" : "";
                }
                GetInstance().response = await GetInstance().GetHttpClient().GetAsync(url);
                GetInstance().response.EnsureSuccessStatusCode();
            }
            catch (Exception e)
            {
                throw new Exception("The request to the API failed.", e);
            }

            dynamic decodedResponse = await instance.DecodeResponse();
            return decodedResponse;
        }

        /// <summary>
        /// Performs a post http request to Fabplan
        /// </summary>
        /// <param name="url"> The url</param>
        /// <param name="data">The body of the post request</param>
        /// <returns>The decoded response</returns>
        public static async Task<dynamic> Post(string url, object data)
        {
            try
            {
                StringContent httpPostBody = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                instance.response = await GetInstance().GetHttpClient().PostAsync(url, httpPostBody);
                instance.response.EnsureSuccessStatusCode();
            }
            catch (Exception e)
            {
                throw new Exception("The request to the API failed.", e);
            }

            dynamic decodedResponse = await instance.DecodeResponse();
            return decodedResponse;
        }


        /// <summary>
        /// Decodes the response from Fabplan
        /// </summary>
        /// <exception cref="UnexpectedFabplanHttpResponseFormatException">
        /// Thrown when the response from Fabplan does not respect the expected response format 
        /// {"status":status, "success": {"data": successData}, "failure": {"message": failureMessage}}.
        /// </exception>
        /// <exception cref="FabplanHttpResponseWarningException">
        /// Thrown when the status of the response is set to warning, meaning that there was a warning issued by the server, but the server 
        /// was still able to send a properly formatted response.
        /// <exception cref="FabplanHttpResponseFailureException">
        /// Thrown when the status of the response is set to failure, meaning that there was an error on the server's side, but the server 
        /// was still able to send a properly formatted response.
        /// </exception>
        /// <returns>The decoded response</returns>
        private async Task<dynamic> DecodeResponse()
        {
            dynamic responseObject;
            try
            {
                responseObject = JsonConvert.DeserializeObject(await instance.response.Content.ReadAsStringAsync());
            }
            catch (Exception e)
            {
                throw new UnexpectedFabplanHttpResponseFormatException("Response is not an object.", e);
            }

            string status;
            try
            {
                status = responseObject.GetValue("status");
            }
            catch (Exception e)
            {
                throw new UnexpectedFabplanHttpResponseFormatException("The \"status\" member of the response is missing.", e);
            }

            if (status == "success")
            {
                dynamic successMember;
                try
                {
                    successMember = responseObject.GetValue("success");
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException("The \"success\" member of the response is missing.", e);
                }

                try
                {
                    return successMember.GetValue("data");
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException(
                        "The \"data\" member of the \"success\" member of the response is missing.",
                        e
                    );
                }
            }
            else if (status == "warning")
            {
                dynamic warningMember;
                try
                {
                    warningMember = responseObject.GetValue("warning");
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException("The \"warning\" member of the response is missing.", e);
                }

                try
                {
                    string message = warningMember.GetValue("message");
                    throw new FabplanHttpResponseWarningException(message);
                }
                catch (FabplanHttpResponseWarningException e)
                {
                    throw e;
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException(
                        "The \"message\" member of the \"warning\" member of the response is missing.",
                        e
                    );
                }
            }
            else if (status == "failure")
            {
                dynamic failureMember;
                try
                {
                    failureMember = responseObject.GetValue("failure");
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException("The \"failure\" member of the response is missing.", e);
                }

                try
                {
                    string message = failureMember.GetValue("message");
                    throw new FabplanHttpResponseFailureException(message);
                }
                catch (FabplanHttpResponseFailureException e)
                {
                    throw e;
                }
                catch (Exception e)
                {
                    throw new UnexpectedFabplanHttpResponseFormatException(
                        "The \"message\" member of the \"failure\" member of the response is missing.",
                        e
                    );
                }
            }
            else
            {
                throw new UnexpectedFabplanHttpResponseFormatException("Inner response status is invalid.");
            }
        }
    }

    /// <summary>
    /// An exception thrown when the response returned by Fabplan is not formatted in the expected fashion. 
    /// </summary>
    public sealed class UnexpectedFabplanHttpResponseFormatException : Exception
    {
        public UnexpectedFabplanHttpResponseFormatException(string message, Exception innerException = null) :
            base(message, innerException)
        { }
    }

    /// <summary>
    /// An exception thrown when there is a server side error from Fabplan. 
    /// </summary>
    public sealed class FabplanHttpResponseFailureException : Exception
    {
        public FabplanHttpResponseFailureException(string message) : base(message)
        { }
    }

    /// <summary>
    /// An exception thrown when there is a server side warning from Fabplan. 
    /// </summary>
    public sealed class FabplanHttpResponseWarningException : Exception
    {
        public FabplanHttpResponseWarningException(string message) : base(message)
        { }
    }
}
