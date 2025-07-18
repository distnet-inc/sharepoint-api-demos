# sharepoint-api-demos
Les scripts fournis dans ce dépôt sont des exemples d'utilisation de Microsoft Graph API pour interagir avec SharePoint Online. Ils permettent de rechercher et télécharger des fichiers basés sur des métadonnées spécifiques, comme une date de rapport.

L'utilisation de Microsoft Graph API permet une approche moderne et sécurisée pour accéder aux ressources SharePoint dans une grande variété de languages de programmation.

Ces scripts ne sont qu’un exemple d’interaction avec SharePoint. D’autres méthodes ou SDKs peuvent également être utilisés, selon le contexte et le langage que vous utilisez.

---


## :ballot_box_with_check: Étapes générales

1. **Installer le certificat SSL qui a été fourni par Distnet**
   
   Le certificat est nécessaire pour établir une connexion sécurisée avec l’API Microsoft Graph. Il doit être installé sur la machine qui exécutera le script.
 
   Pour faire l'installation sur un ordinateur ou serveur Windows, vous pouvez exécuter le script PowerShell suivant:
   > Remplacer le chemin du fichier PFX par le chemin réel où le certificat est stocké.  
   > Le mot de passe du certificat vous sera fourni par Distnet
   ```powershell
   $pfxPassword = Read-Host -AsSecureString "Enter PFX password"

   Import-PfxCertificate -FilePath "C:\path\to\cert.pfx" `
       -CertStoreLocation "Cert:\CurrentUser\My" `
       -Password $pfxPassword `
       -Exportable
   ```

   Copiez l’empreinte numérique (thumbprint) du certificat installé, elle sera nécessaire pour l’étape suivante.

1. **Établir une connexion sécurisée à Microsoft Graph**

   Les autorisations configurées pour accéder à l’API sont de type "Application", ce qui permet d’exécuter des actions sans interaction utilisateur. Cela nécessite un certificat valide pour l'authentification.

2. **Résoudre le site SharePoint qui vous a été attribué**

   Identifier l’ID interne du site SharePoint à partir du nom du site.

3. **Localiser la librairie de documents**

   Trouver l’identifiant de la librairie (document library) à partir de son nom affiché (display name).

4. **Rechercher les fichiers selon les métadonnées**

   Interroger les fichiers présents dans la librairie, filtrés par une valeur des métadonnées.
   Deux métadonnées personnalisées peuvent être interrogées :
   - `ReportDate` : Date du rapport, au format ISO 8601 (YYYY-MM-DD).
   - `ReportType` : Nom du type de rapport. Les valeurs possible sont:

      | Français                                      | Anglais                                     |
      |-----------------------------------------------|---------------------------------------------|
      | `Âge des comptes`                             | `Receivables aging`                         |
      | `Âge des comptes - Excel`                     | `Receivables aging - Excel`                 |
      | `Encaissements par débiteurs`                 | `Receipts List`                             |
      | `Encaissements par débiteurs - Excel`         | `Invoice payments by debtors - Excel`       |
      | `Factures impayées`                           | `Outstanding Invoices`                      |
      | `Factures impayées - Excel`                   | `Outstanding Invoices - Excel`              |
      | `Factures inéligibles`                        | `Ineligible Invoices`                       |
      | `Factures inéligibles - Excel`                | `Ineligible Invoices - Excel`               |
      | `Rapport journalier`                          | `Daily Report`                              |
      | `Reconnaissance de dettes non-reçues`         | `Debt Acknowledgments not received`         |
      | `Reconnaissance de dettes non-reçues - Excel` | `Debt Acknowledgments not received - Excel` |
      | `Sommaire`                                    | `Statement`                                 |
      | `Sommaire - Excel`                            | `Statement - Excel`                         |

5. **Télécharger les fichiers trouvés**

   Télécharger chaque fichier correspondant aux critères, en enregistrant une copie locale dans le dossier de sortie, tout en récupérant au besoin des métadonnées supplémentaires.