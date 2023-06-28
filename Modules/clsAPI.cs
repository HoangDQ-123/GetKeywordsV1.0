using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using OpenAI_API;

using Newtonsoft.Json;


namespace GetKeywords.Modules
{
    public class clsAPI
    {
        public static async Task<string> CallChatGPTAPI(string strContent)
        {
            HttpClient client = new HttpClient();
            
            // Nạp KeyChatGPT vào Header này.
            client.DefaultRequestHeaders.Add("authorization", "Bearer " + InitVar.v_arrKeyChatGPT[0]); // điền keyChatGPT

            // Thay đổi các tham số tùy biến vào sau khi đã lấy được các biến Config ChatGPT
            var content = new StringContent("{\"model\": \"text-davinci-003\", \"prompt\": \"" + strContent + "\",\"temperature\": 0,\"max_tokens\": 4000}", Encoding.UTF8, "application/json");
            
            HttpResponseMessage response = await client.PostAsync("https://api.openai.com/v1/completions", content);

            string responseString = await response.Content.ReadAsStringAsync();
            try
            {
                var dyData = JsonConvert.DeserializeObject<dynamic>(responseString);
                string guess = "";
                if (dyData != null)
                {
                    guess = dyData.choices[0].text;
                }
                   
                return guess;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"---> Could not deserialize the JSON: {ex.Message}");
                return "";
            }
        }
    }
}
