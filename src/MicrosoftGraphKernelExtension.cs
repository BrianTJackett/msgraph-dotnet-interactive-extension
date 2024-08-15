// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.DotNet.Interactive.Commands;
using Microsoft.DotNet.Interactive.CSharp;
using Microsoft.DotNet.Interactive.Directives;
using Microsoft.Graph;
using Beta = Microsoft.Graph.Beta;

namespace Microsoft.DotNet.Interactive.MicrosoftGraph;

/// <summary>
/// .NET Interactive magic command extension to provide
/// authenticated Microsoft Graph clients.
/// </summary>
public class MicrosoftGraphKernelExtension : IKernelExtension
{
    /// <summary>
    /// Main entry point to extension, invoked via
    /// "#!microsoftgraph".
    /// </summary>
    /// <param name="kernel">The .NET Interactive kernel the extension is loading into.</param>
    /// <returns>A completed System.Task.</returns>
    public Task OnLoadAsync(Kernel kernel)
    {
        if (kernel is not CompositeKernel cs)
        {
            return Task.CompletedTask;
        }

        var cSharpKernel = cs.ChildKernels.OfType<CSharpKernel>().FirstOrDefault();
        if (cSharpKernel == null)
        {
            return Task.CompletedTask;
        }

        KernelDirectiveParameter clientIdOption = new(
            "--client-id",
            description: "Application (client) ID registered in Azure Active Directory.");

        KernelDirectiveParameter tenantIdOption = new(
            "--tenant-id",
            description: "Directory (tenant) ID in Azure Active Directory.");
            //getDefaultValue: () => "common");

        KernelDirectiveParameter clientSecretOption = new(
            "--client-secret",
            description: "Application (client) secret registered in Azure Active Directory.");

        KernelDirectiveParameter configFileOption = new(
            "--config-file",
            description: "JSON file containing any combination of tenant ID, client ID, and client secret. Values are only used if corresponding option is not passed to the magic command.");

        KernelDirectiveParameter scopeNameOption = new(
            "--scope-name",
            description: "Scope name for Microsoft Graph connection.");
        //getDefaultValue: () => "graphClient");

        KernelDirectiveParameter authenticationFlowOption = new(
            "--authentication-flow",
            description: "Azure Active Directory authentication flow to use.");
        //getDefaultValue: () => AuthenticationFlow.InteractiveBrowser);

        KernelDirectiveParameter nationalCloudOption = new KernelDirectiveParameter(
            "--national-cloud",
            description: "National cloud for authentication and Microsoft Graph service root endpoint.").AddCompletions(_ => Enum.GetValues<NationalCloud>().Select(c => c.ToString()));
            //getDefaultValue: () => NationalCloud.Global);

        KernelDirectiveParameter apiVersionOption = new(
            "--api-version",
            description: "Microsoft Graph API version.");
            //getDefaultValue: () => ApiVersion.V1);

        KernelActionDirective graphCommand = new KernelActionDirective("#!microsoftgraph")
        {
            Parameters = [
                clientIdOption,
                tenantIdOption,
                clientSecretOption,
                configFileOption,
                scopeNameOption,
                authenticationFlowOption,
                nationalCloudOption,
                apiVersionOption,
            ],
            Description = "Send Microsoft Graph requests using the specified permission flow."
        };

        graphCommand.SetHandler(
            async (CredentialOptions credentialOptions, string scopeName, AuthenticationFlow authenticationFlow, NationalCloud nationalCloud, ApiVersion apiVersion) =>
            {
                try
                {
                    credentialOptions.ValidateOptionsForFlow(authenticationFlow);
                }
                catch (AggregateException ex)
                {
                    KernelInvocationContextExtensions.DisplayStandardError(
                        KernelInvocationContext.Current,
                        $"INVALID INPUT: {ex.Message}");
                    return;
                }

                var tokenCredential = CredentialProvider.GetTokenCredential(
                    authenticationFlow, credentialOptions, nationalCloud);

                switch (apiVersion)
                {
                    case ApiVersion.V1:
                        GraphServiceClient graphServiceClient = new(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClient.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClient, typeof(GraphServiceClient));
                        break;
                    case ApiVersion.Beta:
                        Beta.GraphServiceClient graphServiceClientBeta = new(tokenCredential, Scopes.GetScopes(nationalCloud));
                        graphServiceClientBeta.RequestAdapter.BaseUrl = BaseUrl.GetBaseUrl(nationalCloud, apiVersion);
                        await cSharpKernel.SetValueAsync(scopeName, graphServiceClientBeta, typeof(Beta.GraphServiceClient));
                        break;
                    default:
                        break;
                }

                KernelInvocationContextExtensions.Display(KernelInvocationContext.Current, $"Graph client declared with name: {scopeName}");
            },
            new CredentialOptionsBinder(clientIdOption, tenantIdOption, clientSecretOption, configFileOption),
            scopeNameOption,
            authenticationFlowOption,
            nationalCloudOption,
            apiVersionOption);

        cSharpKernel.AddDirective(graphCommand);

        cSharpKernel.DeferCommand(new SubmitCode("using Microsoft.Graph;"));

        return Task.CompletedTask;
    }
}
