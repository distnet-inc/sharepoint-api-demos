using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Security.Cryptography.X509Certificates;

class Program
{
    static async Task Main()
    {
        // === Config ===
        string tenantId = "{TenantID}";
        string clientId = "{ClientId}";
        string certificateThumbprint = "{Thumbprint du certificat}";
        string siteHostname = "{Adresse SharePoint}";
        string sitePath = "{Nom du site Sharepoint}"; // Nom du site Sharepoint (celui qu'on retrouve dans l'URL)
        string libraryName = "{Nom de la librairie}"; // Le DisplayName de la librairie Sharepoint (Par défaut "Rapports")
        string reportDate = "{Date du rapport sous format yyyy-MM-dd}"; // Exemple: "2025-07-17"
        string outputFolder = @"C:\Temp";

        Console.WriteLine("Connecting to Microsoft Graph...");

        // === Load Certificate ===
        var cert = GetCertificateFromStore(certificateThumbprint);
        if (cert == null)
        {
            Console.WriteLine("Certificate not found.");
            return;
        }

        var authProvider = new ClientCertificateCredential(tenantId, clientId, cert);
        var graphClient = new GraphServiceClient(authProvider);

        Console.WriteLine("Connected to Microsoft Graph.");

        // === Resolve Site ===
        Console.WriteLine($"Resolving site '{sitePath}'...");
        var site = await graphClient
            .Sites[$"{siteHostname}:/sites/{sitePath}"]
            .GetAsync();

        if (site == null)
        {
            Console.WriteLine("Site not found.");
            return;
        }

        Console.WriteLine($"Site resolved: {site.Id}");

        // === Get Document Library ===
        Console.WriteLine($"Getting document library '{libraryName}'...");
        var lists = await graphClient
            .Sites[site.Id]
            .Lists
            .GetAsync();

        var list = lists?.Value?.FirstOrDefault(l => l.DisplayName == libraryName);
        if (list == null)
        {
            Console.WriteLine($"Document library '{libraryName}' not found.");
            return;
        }
        Console.WriteLine($"Library found: {list.Id}");

        // === Query Files by Metadata ===
        Console.WriteLine($"Searching files with ReportDate = '{reportDate}'...");

        // Construct the URL: /sites/{siteId}/lists/{listId}/items?$filter=fields/ReportDate eq '2024-07-01'&$expand=driveItem
        var requestInfo = graphClient.Sites[site.Id].Lists[list.Id].Items.ToGetRequestInformation(requestConfig =>
        {
            requestConfig.QueryParameters.Filter = $"fields/ReportDate eq '{reportDate}'";
            requestConfig.QueryParameters.Expand = new[] { "driveItem" };
        });

        var response = await graphClient.RequestAdapter.SendAsync(
            requestInfo,
            ListItemCollectionResponse.CreateFromDiscriminatorValue,
            default
        );
        var items = response?.Value;

        if (items == null || items.Count == 0)
        {
            Console.WriteLine($"No files found with ReportDate = '{reportDate}'.");
            return;
        }
        Console.WriteLine($"Found {items.Count} file(s).");

        // === Prepare Output Directory ===
        if (!Directory.Exists(outputFolder))
        {
            Console.WriteLine($"Creating output folder: {outputFolder}");
            Directory.CreateDirectory(outputFolder);
        }

        // === Download Files ===
        Console.WriteLine($"Starting file download...");
        int counter = 0;

        foreach (var item in items)
        {
            var driveItem = item.DriveItem;

            if (driveItem == null || driveItem.ParentReference?.DriveId == null || driveItem.Id == null || driveItem.File == null || driveItem.Name == null)
            {
                Console.WriteLine($"Skipping file '{driveItem?.Name}' due to missing DriveId, ItemId, File or Name.");
                continue;
            }

            string name = driveItem.Name;
            string driveId = driveItem.ParentReference.DriveId;
            string itemId = driveItem.Id;

            string reportType = "Unknown";
            if (item.Fields?.AdditionalData != null && item.Fields.AdditionalData.ContainsKey("ReportType"))
                reportType =  (string)item.Fields.AdditionalData["ReportType"];

            string outPath = Path.Combine(outputFolder, name);

            var stream = await graphClient
                .Drives[driveId]
                .Items[itemId]
                .Content
                .GetAsync();

            if (stream == null)
            {
                Console.WriteLine($"Failed to download file '{name}' (DriveId: {driveId}, ItemId: {itemId}).");
                continue;
            }

            using (var fileStream = File.Create(outPath))
                await stream.CopyToAsync(fileStream);
            
            Console.WriteLine($"Downloaded: [{reportType}] Filename: '{name}'");
            counter++;
        }

        Console.WriteLine($"Completed.");
        Console.WriteLine($"Successfully downloaded {counter} of {items.Count} file(s) to: {outputFolder}");
    }

    static X509Certificate2? GetCertificateFromStore(string thumbprint)
    {
        using (var store = new X509Store(StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadOnly);
            var certs = store.Certificates
                .Find(X509FindType.FindByThumbprint, thumbprint, validOnly: false);
            return certs.Count > 0 ? certs[0] : null;
        }
    }
}
