# Exemple PowerShell pour l'utilisation de Microsoft Graph API avec SharePoint Online

Ce script est un exemple de la façon dont vous pouvez interagir avec SharePoint Online en utilisant Microsoft Graph API et PowerShell.

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

- Installer les modules PowerShell requis
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
  ```

---
## :gear: Configuration du script
Avant de lancer le script, vous devez remplir les valeurs suivantes :

```powershell
  TenantId          = "{TenantID}"
  ClientId          = "{ClientId}"
  CertificateThumb  = "{Thumbprint du certificat}"
  SiteHostname      = "{Adresse SharePoint}"
  SiteName          = "{Nom du site Sharepoint}" # Nom du site Sharepoint (celui qu'on retrouve dans l'URL)
  LibraryName       = "{Nom de la librairie}" # Le DisplayName de la librairie Sharepoint (Par défaut "Rapports")
  ReportDate        = "{Date du rapport sous format yyyy-MM-dd}" # Exemple: "2025-07-17"
  OutputFolder      = "C:\Temp"
```

---

## :rocket: Étapes pour télécharger les rapports
1. **Obtenir le certificat**

   Le certificat est récupéré à partir du magasin de certificats local en utilisant son empreinte numérique (thumbprint) :
   ```powershell
     Get-ChildItem -Path Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $config.CertificateThumb }
   ```

2. **Connexion à Microsoft Graph**

   Le script utilise Connect-MgGraph avec le certificat pour s’authentifier :
   ```powershell
     Connect-MgGraph -ClientId $config.ClientId -TenantId $config.TenantId -Certificate $certificate
   ```
   Cela permet une authentification sécurisée et automatisée.

3. **Résolution du site SharePoint**

   Le site est identifié via le nom d’hôte (SiteHostname) et le nom du site SharePoint (SiteName) :
   ```powershell
     Get-MgSite -SiteId "$($config.SiteHostname):/sites/$($config.SiteName)"  
   ```

   Cette commande retourne l’ID unique du site, nécessaire pour les étapes suivantes.

4. **Récupération de la bibliothèque de documents**
 
   Le site est interrogée pour obtenir la bibliothèque de documents (document library) en utilisant son nom affiché (DisplayName) :

   ```powershell
     Get-MgSiteList -SiteId $site.Id | Where-Object { $_.DisplayName -eq $config.LibraryName }
   ```

5. Recherche de fichiers selon une date et/ou un type de rapport

   Le script interroge les fichiers dans la bibliothèque en utilisant un filtre basé sur une métadonnée personnalisée (cette exemple retourne tous les rapports d'une certaine date).
 
   ```powershell
     Get-MgSiteListItem -SiteId $site.Id -ListId $list.Id `
        -Filter "fields/ReportDate eq '$($config.ReportDate)'" -ExpandProperty "driveItem"
    ```

6. Téléchargement des fichiers

   Les fichiers sont téléchargés dans un dossier à l’aide de :

   ```powershell
     Get-MgDriveItemContent -DriveId $driveId -DriveItemId $itemId -OutFile $outPath
   ```

   > Pour récupérer le type de rapport, vous pouvez y accéder par
   > ```powershell
   >  $file.Fields.AdditionalProperties["ReportType"]
   > ```