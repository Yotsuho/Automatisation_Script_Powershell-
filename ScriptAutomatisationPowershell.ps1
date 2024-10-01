# Menu principal
do {
    Clear-Host
    Write-Host "Menu Principal"
    Write-Host "1. Créer une OU"
    Write-Host "2. Créer un utilisateur"
    Write-Host "3. Créer un groupe"
    Write-Host "4. Importer des utilisateurs via CSV"
    Write-Host "5. Exporter un fichier CSV"
    Write-Host "6. Sauvegarde de L'AD"
    Write-Host "7. Restauration de L'AD"
    Write-Host "8. Quitter"

    # Récupérer le choix de l'utilisateur
    $choix = Read-Host "Faites votre choix (1, 2, 3, 4, 5, 6, 7, 8)"
    
    switch ($choix) {
        1 {
            # Option pour créer une OU
            $nomOU = Read-Host "Entrez le nom de l'OU"
            New-ADOrganizationalUnit -Name $nomOU -Path "DC=NDS,DC=local"
            Write-Host "OU '$nomOU' créée avec succès."
        }
        2 {
            # Option pour créer un utilisateur
            $nomUtilisateur = Read-Host "Entrez le nom de l'utilisateur"
            $motDePasse = Read-Host "Entrez le mot de passe" -AsSecureString
            $description = Read-Host "Entrez la description de l'utilisateur"
            
            # Créer l'utilisateur
            New-ADUser -SamAccountName $nomUtilisateur -UserPrincipalName "$nomUtilisateur@NDS.local" -Name $nomUtilisateur -GivenName $nomUtilisateur -Surname "Nom" -Enabled $true -AccountPassword $motDePasse -Description $description -PassThru | Set-ADUser
            
            Write-Host "Utilisateur '$nomUtilisateur' créé avec succès."

            # Demander si l'utilisateur veut ajouter l'utilisateur à un groupe
            $ajouterAuGroupe = Read-Host "Voulez-vous ajouter l'utilisateur à un groupe ? (Oui/Non)"
            if ($ajouterAuGroupe -eq "Oui") {
                $nomGroupe = Read-Host "Entrez le nom du groupe"
                Add-ADGroupMember -Identity $nomGroupe -Members $nomUtilisateur
                Write-Host "Utilisateur '$nomUtilisateur' ajouté au groupe '$nomGroupe'."
                
                # Créer le répertoire partagé personnel
                Install-Module NTFSSecurity
                $repertoire = "C:\UsersPartage\$nomGroupe"

                if (Test-Path $repertoire -PathType Container) {
                    # Si le répertoire existe, ajouter simplement l'utilisateur au partage
                    Add-NTFSAccess -Path $repertoire -Account "$nomUtilisateur@NDS.local" -AccessRights Modify
                    Write-Host "L'utilisateur '$nomUtilisateur' a été ajouté au répertoire partagé existant '$nomGroupe'."
                    Grant-SmbShareAccess -Name $nomGroupe -AccountName "$nomUtilisateur@NDS.local" -AccessRight Change
                }
                else{
                    # Si le répertoire n'existe pas, le créer, définir les autorisations NTFS et partager le dossier
                    New-Item -ItemType Directory -Path $repertoire
                    Write-Host "Répertoire partagé personnel créé pour '$nomUtilisateur'."

                    # Ajout des autorisations NTFS
                    Add-NTFSAccess -Path $repertoire -Account "$nomUtilisateur@NDS.local" -AccessRights Modify

                    # Partage du dossier
                    New-SmbShare -Name $nomGroupe -Path $repertoire -ChangeAccess "$nomUtilisateur@NDS.local"
                    Grant-SmbShareAccess -Name $nomGroupe -AccountName "$nomUtilisateur@NDS.local" -AccessRight Change
                    Set-ADUser $nomUtilisateur -HomeDrive "U:" -HomeDirectory "\\Winservnds\$nomGroupe" -Enabled $true
                }

                #New-Item -ItemType Directory -Path "C:\UsersPartage\$nomGroupe"
                #Write-Host "Répertoire partagé personnel créé pour '$nomUtilisateur'."
                
                #Ajout des autorisations NTFS
                #Add-NTFSAccess -Path "C:\UsersPartage\$nomGroupe" -Account "$nomUtilisateur@NDS.local" -AccessRights Modify

                #Partage du dossier
                #New-SmbShare -Name $nomGroupe -Path "C:\UsersPartage\$nomGroupe" -ChangeAccess "$nomUtilisateur@NDS.local" 
            }
        }
        3 {
            # Option pour créer un groupe
            $nomGroupe = Read-Host "Entrez le nom du groupe"
            $descriptionGroupe = Read-Host "Entrez la description du groupe"

            $ouGroupe = Read-Host "Entrez le chemin complet de l'OU pour le groupe ( OU=NomOU,DC=NDS,DC=local ).Si vide, le groupe sera créé à la racine"
            if (-not $ouGroupe) {
                $ouGroupe = "DC=NDS,DC=local"
            }

            New-ADGroup -Name $nomGroupe -GroupScope Global -GroupCategory Security -Description $descriptionGroupe -Path $ouGroupe
            Write-Host "Groupe '$nomGroupe' créé avec succès dans l'OU '$ouGroupe'."

        }
        4 {
            # Chemin du fichier CSV contenant les utilisateurs
            $cheminFichierCSV = "C:\Users\Administrateur\caca.csv"

            # Importer les données du fichier CSV
            $utilisateurs = Import-Csv -Path $cheminFichierCSV

            # Parcourir chaque utilisateur dans le fichier CSV
            foreach ($utilisateur in $utilisateurs) {
                # Récupérer les données de l'utilisateur
                $nomUtilisateur = $utilisateur.NomUtilisateur
                $prenomUtilisateur = $utilisateur.PrenomUtilisateur
                $motDePasse = ConvertTo-SecureString $utilisateur.MotDePasse -AsPlainText -Force
                $description = $utilisateur.Description

                # Créer l'utilisateur
                New-ADUser -SamAccountName $nomUtilisateur -UserPrincipalName "$nomUtilisateur@mondomaine.com" -Name "$prenomUtilisateur $nomUtilisateur" -GivenName $prenomUtilisateur -Surname $nomUtilisateur -Enabled $true -AccountPassword $motDePasse -Description $description

                Write-Host "Utilisateur '$prenomUtilisateur $nomUtilisateur' créé avec succès."

                # Vérifier si des groupes sont spécifiés pour l'utilisateur dans le fichier CSV
                if ($utilisateur.Groupes) {
                    $groupes = $utilisateur.Groupes -split "," | ForEach-Object { $_.Trim() }

                    # Ajouter l'utilisateur à chaque groupe spécifié
                    foreach ($groupe in $groupes) {
                        Add-ADGroupMember -Identity $groupe -Members $nomUtilisateur
                        Write-Host "Utilisateur '$prenomUtilisateur $nomUtilisateur' ajouté au groupe '$groupe'."
                    }  
                }
            }
        }

        5 {
            # Nom du fichier CSV exporté
            $nomFichierExport = Read-Host "Entrez le nom du fichier"

            # Spécifier le chemin de destination pour le fichier CSV
            $cheminDestinationCSV = "C:\Users\Administrateur\$nomFichierExport"

            # Obtenir tous les utilisateurs d'Active Directory
            $utilisateurs = Get-ADUser -Filter *

            # Exporter les utilisateurs vers un fichier CSV
            $utilisateurs | Select-Object SamAccountName, GivenName, Surname, Enabled, Description | Export-Csv -Path $cheminDestinationCSV -NoTypeInformation

            Write-Host "Exportation des utilisateurs terminée. Le fichier CSV a été enregistré à l'emplacement : $cheminDestinationCSV"
        }

        6 {  
            try{
                $Path = "E:\backup\ActiveDirectory"
                # create backup file name
                $Filename = "ADBackupFull" + "#" + $((Get-Date -Format s) -replace ":","-") + ".bak"
                $Filepath = Join-Path $Path $Filename

                # backup active directory
                Invoke-Expression 'ntdsutil "activate instance ntds" ifm "create full $Filepath" quit quit'
    
                # get dates for backup retention exclusion
                $Today = Get-Date -Format d
                $FirstDateOfWeek = Get-Date (Get-Date).AddDays(-[int](Get-Date).Dayofweek) -Format d
                $FirstDateOfMonth = Get-Date -Day 1 -Format d

                # delete all backups except for today, first day of week and first day of month
                Get-ChildItem $Path | select *,@{L="CreationTimeDate";E={Get-Date $_.CreationTime -Format d}} | Group-Object CreationTimeDate | %{
        
                    # only one backup per day
                    if($_.Count -gt 1){
            
                        $_.Group | Sort-Object CreationTime -Descending | Select-Object -Skip 1     
                    }
                
                    # keep only required backups
                    $_.Group | Where-Object{$_.CreationTimeDate -ne $Today -and $_.CreationTimeDate -ne $FirstDateOfWeek -and $_.CreationTimeDate -ne $FirstDateOfMonth}
            
                } | Remove-Item -Recurse -Force
    
                }catch{

                    Write-PPErrorEventLog -Source "Backup ActiveDirectory" -ClearErrorVariable
                }
           }
        
        7 {
            try {
                # Chemin vers le dossier de sauvegarde
                $cheminSauvegarde = "E:\backup\ActiveDirectory\ADBackupFull#2024-02-06T12-29-44.bak"

                # Chemin vers les fichiers de sauvegarde
                $cheminFichiersSauvegarde = Join-Path $cheminSauvegarde "Active directory"

                # Désactiver le service Active Directory Domain Services
                Stop-Service -Name "NTDS" -Force

                # Arrêter le service de Registre
                Stop-Service -Name "RemoteRegistry" -Force

                # Restaurer les fichiers de sauvegarde
                Copy-Item -Path $cheminFichiersSauvegarde\ntds.dit -Destination "$env:SystemRoot\NTDS" -Force
                Copy-Item -Path $cheminFichiersSauvegarde\ntds.jfm -Destination "$env:SystemRoot\NTDS" -Force
                Copy-Item -Path $cheminSauvegarde\registry\SYSTEM -Destination "$env:SystemRoot\System32\config" -Force
                Copy-Item -Path $cheminSauvegarde\registry\SECURITY -Destination "$env:SystemRoot\System32\config" -Force

                # Activer à nouveau le service Active Directory Domain Services
                Start-Service -Name "NTDS"

                # Démarrer à nouveau le service de Registre
                Start-Service -Name "RemoteRegistry"    

                Write-Host "Restauration de la sauvegarde de l'Active Directory terminée avec succès."
            }
            catch {
                Write-Host "Une erreur s'est produite lors de la restauration de la sauvegarde de l'Active Directory : $_"
            }

        }

        8 {
            # Option pour quitter
            Write-Host "Au revoir !"
        }
        default {
            Write-Host "Choix invalide. Veuillez entrer 1, 2, 3, 4, 5, 6, 7, 8."
        }
    }
    
    # Attendre une touche pour revenir au menu principal
    Read-Host "Appuyez sur Entrée pour revenir au menu principal"
} while ($choix -ne 8)
