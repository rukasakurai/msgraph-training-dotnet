// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <SettingsSnippet>
using Microsoft.Extensions.Configuration;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;

public class Settings
{
    public string? ClientId { get; set; }
    public string? ClientSecret { get; set; }
    public string? TenantId { get; set; }
    public string? AuthTenant { get; set; }
    public string[]? GraphUserScopes { get; set; }
    public string? KeyVaultName {get; set; }

    public static Settings LoadSettings()
    {
        // Load settings
        IConfiguration config = new ConfigurationBuilder()
            // appsettings.json is required
            .AddJsonFile("appsettings.json", optional: false)
            // appsettings.Development.json" is optional, values override appsettings.json
            .AddJsonFile($"appsettings.local.json", optional: true)
            // User secrets are optional, values override both JSON files
            .AddUserSecrets<Program>()
            .Build();

        Settings settings = config.GetRequiredSection("Settings").Get<Settings>();

        string kvUri = "https://" + settings.KeyVaultName + ".vault.azure.net";
        var secretClient = new SecretClient(new Uri(kvUri), new DefaultAzureCredential());
        KeyVaultSecret secret = secretClient.GetSecret("clientSecretForGraph");
        settings.ClientSecret = secret.Value;

        return settings;
    }
}
// </SettingsSnippet>
