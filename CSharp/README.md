# Exemple CSharp pour l'utilisation de Microsoft Graph API avec SharePoint Online

Ce code est un exemple de la façon dont vous pouvez interagir avec SharePoint Online en utilisant Microsoft Graph API et CSharp.

## :ballot_box_with_check: Prérequis
Pour authentifier l'application sans interaction manuelle (authentification "app-only"), vous devez utiliser un certificat X.509.

- Le certificat (PFX) doit être installé dans le magasin **CurrentUser** ou **LocalMachine**, selon votre configuration.  
  Celui-ci vous sera fourni par Distnet. 
  
  Pour se faire, vous pouvez utiliser le script PowerShell suivant pour installer le certificat PFX dans le magasin **CurrentUser**:
  ```powershell
   $pfxPassword = Read-Host -AsSecureString "Enter PFX password"

   Import-PfxCertificate -FilePath "C:\path\to\cert.pfx" `
       -CertStoreLocation "Cert:\CurrentUser\My" `
       -Password $pfxPassword `
       -Exportable
   ```
  
  Copiez l’empreinte numérique (thumbprint) du certificat installé, elle sera nécessaire pour l’étape suivante.

- Installer les modules NuGet requis
  
  Les packages NuGet suivants sont nécessaires pour interagir avec Microsoft Graph API :
  `Azure.Identity` et `Microsoft.Graph`.

---

## :gear: Configuration
Avant de lancer le code, vous devez remplir les valeurs suivantes :

```powershell
  string tenantId = "{TenantID}";
  string clientId = "{ClientId}";
  string certificateThumbprint = "{Thumbprint du certificat}";
  string siteHostname = "{Adresse SharePoint}";
  string sitePath = "{Nom du site Sharepoint}"; // Nom du site Sharepoint (celui qu'on retrouve dans l'URL)
  string libraryName = "{Nom de la librairie}"; // Le DisplayName de la librairie Sharepoint (Par défaut "Rapports")
  string reportDate = "{Date du rapport sous format yyyy-MM-dd}"; // Exemple: "2025-07-17"
  string outputFolder = @"C:\Temp";
```

---

## :rocket: Étapes pour télécharger les rapports
1. **Obtenir le certificat**
   
   Le certificat est récupéré à partir du magasin de certificats local en utilisant son empreinte numérique (thumbprint) :
   ```csharp
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
   ```

2. **Connexion à Microsoft Graph**
   
   Le code utilise ClientCertificateCredential avec le certificat pour s’authentifier :
   ```csharp
    var authProvider = new ClientCertificateCredential(tenantId, clientId, cert);
    var graphClient = new GraphServiceClient(authProvider);
   ```
   Cela permet une authentification sécurisée et automatisée.

3. **Résolution du site SharePoint**
   
   Le site est identifié via le nom d’hôte (siteHostname) et le nom du site SharePoint (siteName) :
   ```csharp
    var site = await graphClient
        .Sites[$"{siteHostname}:/sites/{siteName}"]
        .GetAsync(); 
   ```

   Ceci retourne l’ID unique du site, nécessaire pour les étapes suivantes.

4. **Récupération de la bibliothèque de documents**
   
   Le site est interrogée pour obtenir la bibliothèque de documents (document library) en utilisant son nom affiché (DisplayName) :

   ```csharp
    var lists = await graphClient
        .Sites[site.Id]
        .Lists
        .GetAsync();

    var list = lists?.Value?.FirstOrDefault(l => l.DisplayName == libraryName);
   ```

5. Recherche de fichiers selon une date et/ou un type de rapport
   
   Le script interroge les fichiers dans la bibliothèque en utilisant un filtre basé sur une métadonnée personnalisée (cette exemple retourne tous les rapports d'une certaine date).
 
   ```csharp
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
    ```
    
6. Téléchargement des fichiers
   
   Les fichiers sont téléchargés dans un dossier à l’aide de :

   ```csharp
    var stream = await graphClient
       .Drives[driveId]
       .Items[itemId]
       .Content
       .GetAsync();
   ```

   > [!TIP] 
   > Pour récupérer le type de rapport, vous pouvez y accéder par
   > ```csharp
   >  (string)item.Fields.AdditionalData["ReportType"];
   > ```