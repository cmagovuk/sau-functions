using Newtonsoft.Json;

namespace CMA.SAU.AzureFunctions
{
    class Request
    {
        public string Method { get; set; }
        public dynamic Payload { get; set; }

        public Request(string reqBody)
        {
            dynamic data = JsonConvert.DeserializeObject(reqBody);
            this.Method = data.method;
            this.Payload = data.payload;
        }

        public string GetJSON()
        {
            return JsonConvert.SerializeObject(
                this,
                Formatting.None,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }
            );
        }
    }
}