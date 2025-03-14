using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace LM2ReadandList_Customized.API
{
    public class Api_Core
    {
        //以(DB名稱和模式)取得對應連接字串
        public static string get_connectstring(string db_name, bool test_mode = false)
        {
            string result = "";
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(3);
                    string url = Properties.Settings.Default.ConnectStr_Server;

                    string fullUrl = $"{url}/get_connectstring?db_name={db_name}&test_mode={test_mode}";
                    result = Task.Run(async () =>
                    {
                        string response = await client.GetStringAsync(fullUrl);
                        return response;
                    }).GetAwaiter().GetResult();
                }
            }
            catch
            {

            }
            return result;
        }

        //取得雲端啟用狀態
        public static bool get_status()
        {
            bool result = false;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(3);
                    string url = Properties.Settings.Default.ConnectStr_Server;

                    string fullUrl = $"{url}/get_azure_mode";
                    result = Task.Run(async () =>
                    {
                        string response = await client.GetStringAsync(fullUrl);
                        return response == "True";
                    }).GetAwaiter().GetResult();
                }
            }
            catch
            {
                throw (new Exception("api get_status drop"));
            }
            return result;
        }
    }

}
