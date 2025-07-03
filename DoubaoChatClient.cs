using System.Text.Json;

namespace LLmDocxTrans;

using System;
using System.Net.Http;
using System.Text;
using System.Collections.Generic;

public class DoubaoChatClient : ChatClient
{
    private readonly HttpClient _httpClient;

    public DoubaoChatClient(string apikey)
    {
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(600)
        };
        _httpClient.BaseAddress = new Uri("https://ark.cn-beijing.volces.com/api/v3/chat/completions");
        _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apikey}");
    }

    public string Chat(string userContent)
    {
        // 构建请求内容
        var requestBody = new
        {
            model = "doubao-1-5-pro-32k-250115",
            messages = new List<object>
            {
                new { role = "system", content = "你是一个专业的翻译助手。" },
                new { role = "user", content = userContent }
            }
        };

        var jsonContent = new StringContent(
            JsonSerializer.Serialize(requestBody),
            Encoding.UTF8,
            "application/json");

        // 发送请求
        using var response = _httpClient.PostAsync("", jsonContent).Result;

        // 处理响应
        response.EnsureSuccessStatusCode();
        var responseBody = response.Content.ReadAsStringAsync().Result;
        var responseObject = JsonSerializer.Deserialize<ResponseModel>(responseBody);
        var result = responseObject?.choices?[0]?.message?.content ?? string.Empty;
        return result;
    }
}