using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using WF.BASE.Helper;
using WF.BASE.Models;

namespace WF.BASE.Service
{
    public class BaseService
    {
        public Base.Responce.GetAllRequest GetAllRequest(Base.Request.GetAllRequest entity)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            var client = new RestClient(ConfigurationSettings.AppSettings["BaseUrl"]);
            var request = new RestRequest(Constants.GET_ALL_REQUEST);
            request.Method = Method.POST;
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("access_token", entity.Token);
            request.AddParameter("page", entity.Page);
            try
            {
                var response = client.Execute<Base.Responce.GetAllRequest>(request);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    var objData = JsonConvert.DeserializeObject<Base.Responce.GetAllRequest>(response.Content);
                    return objData;
                }
            }
            catch (Exception ex)
            {
            }
            return null;
        }
    }
}
