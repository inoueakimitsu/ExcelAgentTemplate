using System.Threading.Tasks;

namespace RunAgentClient
{
    using ExcelDna.Integration;
    using ExcelDna.Registration.Utils;
    using Newtonsoft.Json;
    using System.Net.Http;
    using System;
    using System.Linq;
    using ExcelDna.Registration;

    /// <summary>
    /// Excel Add-In のエントリーポイントとなるクラスです。
    /// このクラスは、Excel が Add-In をロードまたはアンロードするときに
    /// 呼び出されるメソッドを提供します。
    /// また、Excel 関数の登録と、オプション引数の扱いについての設定も行います。
    /// </summary>
    public class AddIn : IExcelAddIn
    {
        /// <summary>
        /// Excel が Add-In をロードするときに呼び出されます。
        /// ここで Excel 関数の登録を行います。
        /// </summary>
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        /// <summary>
        /// Excel が Add-In をアンロードするときに呼び出されます。
        /// 現在は何も行いません。
        /// </summary>
        public void AutoClose()
        {
        }

        /// <summary>
        /// Excel 関数の登録と、オプション引数の扱いについての設定を行います。
        /// 空のセルへの参照をオプション引数として扱うように設定します。
        /// </summary>
        public void RegisterFunctions()
        {
            var paramConversionConfig = new ParameterConversionConfiguration()
            // 空のセルへの参照の扱いについてのオプションを指定します。
            .AddParameterConversion(ParameterConversions.GetOptionalConversion(treatEmptyAsMissing: true))
            .AddNullableConversion(treatEmptyAsMissing: true, treatNAErrorAsMissing: true);

            ExcelRegistration.GetExcelFunctions()
            .ProcessParameterConversions(paramConversionConfig)
            .RegisterFunctions();
        }
    }


    public static class RemoteAgentRunner
    {
        /// <summary>
        /// Excel 内の文字列を JSON 文字列に変換します。
        /// このために、特殊文字をエスケープします。
        /// </summary>
        /// <param name="text">変換する Excel 内の文字列</param>
        /// <returns>エスケープされた JSON 文字列</returns>
        private static string ConvertExcelStringToJsonString(string text)
        {
            text = text.Replace("\\", "\\\\");
            text = text.Replace("\"", "\\\"");
            text = text.Replace("\b", "\\b");
            text = text.Replace("\r", "\\r");
            text = text.Replace("\n", "\\n");
            text = text.Replace("\t", "\\t");
            text = text.Replace("\f", "\\f");
            return text;
        }

        /// <summary>
        /// JSON 文字列を Excel 内の文字列に変換します。
        /// このために、エスケープされた特殊文字を元に戻します。
        /// </summary>
        /// <param name="response">変換する JSON 文字列</param>
        /// <returns>エスケープが解除された Excel 内の文字列</returns>
        private static string ConvertJsonStringToExcelString(string response)
        {
            response = response.Replace("\\\\", "\\");
            response = response.Replace("\\\"", "\"");
            response = response.Replace("\\b", "\b");
            response = response.Replace("\\r", "\r");
            response = response.Replace("\\n", "\n");
            response = response.Replace("\\t", "\t");
            response = response.Replace("\\f", "\f");
            return response;
        }

        /// <summary>
        /// 文字列の先頭と末尾のダブルクォートを削除します。
        /// 文字列が null または空の場合、そのままの文字列を返します。
        /// </summary>
        /// <param name="str">ダブルクォートを削除する文字列</param>
        /// <returns>ダブルクォートが削除された文字列</returns>
        private static string RemoveQuotes(string str)
        {
            if (string.IsNullOrEmpty(str)) return str;
            // 先頭と末尾のダブルクォートを削除します。
            if (str.StartsWith("\"")) str = str.Substring(1);
            if (str.EndsWith("\"")) str = str.Substring(0, str.Length - 1);
            return str;
        }


        [ExcelFunction(Name = "RunAgent", Description = "Excel から AI エージェントを非同期に実行します。サーバー URL をオプションで指定できます。")]
        public static object RunAgent(
            [ExcelArgument(Name = "inputMessage", Description = "AI エージェントに送信する入力メッセージ")]
            string inputMessage,
            [ExcelArgument(Name = "serverUrl", Description = "AI エージェントがホストされているサーバーの URL。デフォルトは 'http://localhost:8889/chat'")]
            string serverUrl = "http://localhost:8889/chat",
            [ExcelArgument(Name = "model", Description = "使用する AI モデルの名前。デフォルトは 'gpt-4-turbo-preview'")]
            string model = "gpt-4-turbo-preview")
        {
            return AsyncTaskUtil.RunTask(
                "RunAgent",
                new object[] { inputMessage, serverUrl, model }, async () =>
                {
                    return await RunAgentAsync(inputMessage, serverUrl, model);
                });
        }

        /// <summary>
        /// 非同期に AI エージェントを実行します。
        /// この関数は、指定されたサーバー URL に対して HTTP POST リクエストを送信し、
        /// レスポンスを受け取ります。リクエストのボディには、入力メッセージとモデル名が JSON 形式で含まれます。
        /// レスポンスは、JSON 文字列から Excel 文字列に変換され、先頭と末尾のダブルクォートが削除されます。
        /// </summary>
        /// <param name="inputMessage">AI エージェントに送信する入力メッセージ</param>
        /// <param name="serverUrl">AI エージェントがホストされているサーバーの URL</param>
        /// <param name="model">使用する AI モデルの名前</param>
        /// <returns>AI エージェントからのレスポンスメッセージ。JSON 文字列から Excel 文字列に変換され、先頭と末尾のダブルクォートが削除されます。</returns>
        private static async Task<string> RunAgentAsync(string inputMessage, string serverUrl, string model)
        {
            try
            {
                HttpClient client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(1000 * 60 * 10);
                var postData = JsonConvert.SerializeObject(new { message = ConvertExcelStringToJsonString(inputMessage), model = model });
                var content = new StringContent(postData, System.Text.Encoding.UTF8, "application/json");
                var response = await client.PostAsync(serverUrl, content);
                var responseString = await response.Content.ReadAsStringAsync();
                return RemoveQuotes(ConvertJsonStringToExcelString(responseString));
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message;
            }
        }
    }

}
