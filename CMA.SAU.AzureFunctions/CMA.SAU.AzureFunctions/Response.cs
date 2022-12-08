using Newtonsoft.Json;

namespace CMA.SAU.AzureFunctions
{
    class Response
    {
#pragma warning disable IDE1006
        public bool success { get; set; }
        public dynamic data { get; set; }
        public string error { get; set; }
#pragma warning restore IDE1006

        public Response()
        {
            success = true;
            error = null;
            data = null;
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