namespace TeamsInReview.Services;

using Azure;
using Azure.AI.OpenAI;
using Microsoft.Graph;
using OpenAI.Chat;
using System.Runtime.CompilerServices;

public class AI
{
    AzureOpenAIClient _openAIClient;
    string _deploymentName;

    public AI(string endpoint, string key, string deploymentname)
    {
        _openAIClient = new AzureOpenAIClient( new Uri(endpoint),
                           new AzureKeyCredential(key));

        _deploymentName = deploymentname;

    }

    public string GetOverlayText(string aiPrompt)
    {

        ChatClient chatClient = _openAIClient.GetChatClient(_deploymentName);

        ChatCompletion completion = chatClient.CompleteChat(
        [
            // System messages represent instructions or other guidance about how the assistant should behave
            //new SystemChatMessage("You are a helpful assistant that talks like a pirate."),
            // User messages represent user input, whether historical or the most recen tinput
            //new UserChatMessage("Hi, can you help me?"),
            // Assistant messages in a request represent conversation history for responses
            //new AssistantChatMessage("Arrr! Of course, me hearty! What can I do for ye?"),
            new UserChatMessage(aiPrompt)
        ]);

        return completion.Content[0].Text;


    }
}
