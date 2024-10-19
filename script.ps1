# Définir le chemin du répertoire racine contenant les dossiers de mangas
$rootDirectoryPath = "E:\Sync\a"  # <-- MODIFIEZ CE CHEMIN si nécessaire

# Vérifier si le répertoire racine existe
if (-Not (Test-Path -Path $rootDirectoryPath)) {
    Write-Error "Le répertoire racine spécifié n'existe pas : $rootDirectoryPath"
    exit
}

# Récupérer tous les sous-dossiers (mangas) dans le répertoire racine
$mangaDirectories = Get-ChildItem -Path $rootDirectoryPath -Directory

if ($mangaDirectories.Count -eq 0) {
    Write-Host "Aucun dossier de manga trouvé dans le répertoire racine spécifié."
    exit
}

# Parcourir chaque dossier de manga
foreach ($mangaDir in $mangaDirectories) {
    Write-Host "------------------------------"
    Write-Host "Traitement du manga : $($mangaDir.Name)"
    Write-Host "------------------------------"

    $mangaPath = $mangaDir.FullName

    # Définir le chemin du dossier 'fusionné' à l'intérieur du dossier de manga
    $fusionneDir = Join-Path -Path $mangaPath -ChildPath "fusionné"

    # Créer le dossier 'fusionné' s'il n'existe pas
    if (-Not (Test-Path -Path $fusionneDir)) {
        try {
            New-Item -Path $fusionneDir -ItemType Directory | Out-Null
            Write-Host "Création du dossier 'fusionné' : $fusionneDir"
        }
        catch {
            Write-Error "Impossible de créer le dossier 'fusionné' : $_"
            continue  # Passer au prochain manga si la création échoue
        }
    }
    else {
        Write-Host "Le dossier 'fusionné' existe déjà : $fusionneDir"
    }

    # Récupérer tous les fichiers .cbz dans le dossier de manga
    $cbzFiles = Get-ChildItem -Path $mangaPath -Filter "*.cbz"

    if ($cbzFiles.Count -eq 0) {
        Write-Host "Aucun fichier .cbz trouvé dans le dossier : $mangaPath"
        continue  # Passer au prochain manga
    }

    # Grouper les fichiers par Titre et Numéro de Volume en utilisant une expression régulière
    $groupedFiles = $cbzFiles | Group-Object {
        # Utiliser uniquement le nom du fichier pour le matching
        if ($_.Name -match "^(?<Titre>.+?)\s--\sVolume\s*(?<Volume>\d+)\s-\sChapitre\s*(?<Chapitre>\d+)\s-") {
            # Retourner une combinaison unique de Titre et Volume pour le groupement
            return "$($matches.Titre.Trim())|$($matches.Volume)"
        }
        else {
            return "SansVolume"
        }
    }

    # Traiter chaque groupe (Titre + Volume)
    foreach ($group in $groupedFiles) {
        # Ignorer les fichiers sans numéro de volume si nécessaire
        if ($group.Name -eq "SansVolume") {
            Write-Host "Ignoré : Certains fichiers n'ont pas de numéro de volume dans le dossier '$($mangaDir.Name)'."
            continue
        }

        # Séparer Titre et Volume
        $splitName = $group.Name -split "\|"
        $titre = $splitName[0]
        $volumeNumber = [int]$splitName[1]

        Write-Host "Traitement du Titre : '$titre', Volume : $volumeNumber..."

        # Définir le nom et le chemin du fichier CBZ fusionné
        $mergedCbzName = "$titre -- Volume $($volumeNumber.ToString("D3")) - Fusionné.cbz"
        $mergedCbzPath = Join-Path -Path $fusionneDir -ChildPath $mergedCbzName

        # Vérifier si le fichier fusionné existe déjà
        if (Test-Path -Path $mergedCbzPath) {
            Write-Warning "Le fichier fusionné existe déjà : $mergedCbzPath. Il sera remplacé."
            try {
                Remove-Item -Path $mergedCbzPath -Force
            }
            catch {
                Write-Error "Impossible de supprimer le fichier existant : $mergedCbzPath. Erreur : $_"
                continue  # Passer au prochain groupe si la suppression échoue
            }
        }

        # Créer un répertoire temporaire unique pour la fusion des images
        $tempDir = Join-Path -Path $mangaPath -ChildPath "Temp_Merge_${titre}_Volume_$volumeNumber"
        if (Test-Path -Path $tempDir) {
            Write-Warning "Le répertoire temporaire existe déjà : $tempDir. Suppression en cours."
            try {
                Remove-Item -Path $tempDir -Recurse -Force
            }
            catch {
                Write-Error "Impossible de supprimer le répertoire temporaire existant : $tempDir. Erreur : $_"
                continue
            }
        }

        try {
            New-Item -Path $tempDir -ItemType Directory | Out-Null
        }
        catch {
            Write-Error "Impossible de créer le répertoire temporaire : $tempDir. Erreur : $_"
            continue
        }

        # Créer un répertoire temporaire pour l'extraction individuelle
        $extractTempDir = Join-Path -Path $mangaPath -ChildPath "Extract_Temp_${titre}_Volume_$volumeNumber"
        if (Test-Path -Path $extractTempDir) {
            Write-Warning "Le répertoire d'extraction temporaire existe déjà : $extractTempDir. Suppression en cours."
            try {
                Remove-Item -Path $extractTempDir -Recurse -Force
            }
            catch {
                Write-Error "Impossible de supprimer le répertoire d'extraction temporaire existant : $extractTempDir. Erreur : $_"
                continue
            }
        }

        try {
            New-Item -Path $extractTempDir -ItemType Directory | Out-Null
        }
        catch {
            Write-Error "Impossible de créer le répertoire d'extraction temporaire : $extractTempDir. Erreur : $_"
            continue
        }

        # Initialiser le compteur d'images
        $imageCounter = 1

        # Trier les fichiers CBZ par numéro de chapitre pour maintenir l'ordre
        $sortedGroup = $group.Group | Sort-Object {
            if ($_.Name -match "Chapitre\s*(\d+)") {
                return [int]$matches[1]
            }
            else {
                return 0
            }
        }

        foreach ($file in $sortedGroup) {
            Write-Host "Extraction de $($file.Name)..."
            try {
                # Extraire le fichier CBZ dans le répertoire d'extraction temporaire
                Expand-Archive -Path $file.FullName -DestinationPath $extractTempDir -Force
            }
            catch {
                Write-Error "Erreur lors de l'extraction de $($file.Name) : $_"
                continue
            }

            # Récupérer tous les fichiers d'image extraits, triés par nom
            $imageFiles = Get-ChildItem -Path $extractTempDir -File | Sort-Object Name

            foreach ($image in $imageFiles) {
                # Définir le nouveau nom avec un compteur séquentiel, par exemple 0001.jpg, 0002.png, etc.
                $extension = $image.Extension
                $newName = "{0:D4}{1}" -f $imageCounter, $extension  # Utilise 4 chiffres pour plus de capacité
                $destinationPath = Join-Path -Path $tempDir -ChildPath $newName

                try {
                    # Copier et renommer l'image dans le répertoire de fusion
                    Copy-Item -Path $image.FullName -DestinationPath $destinationPath -Force
                    $imageCounter++
                }
                catch {
                    Write-Error "Erreur lors de la copie/renommage de $($image.Name) : $_"
                }
            }

            # Nettoyer le répertoire d'extraction temporaire pour le prochain fichier CBZ
            try {
                Remove-Item -Path "$extractTempDir\*" -Recurse -Force
            }
            catch {
                Write-Error "Impossible de nettoyer le répertoire d'extraction temporaire : $extractTempDir. Erreur : $_"
            }
        }

        # Supprimer le répertoire d'extraction temporaire après traitement de tous les fichiers CBZ
        try {
            Remove-Item -Path $extractTempDir -Recurse -Force
        }
        catch {
            Write-Error "Impossible de supprimer le répertoire d'extraction temporaire : $extractTempDir. Erreur : $_"
        }

        # Compresser le contenu fusionné en un nouveau fichier CBZ
        Write-Host "Compression en $mergedCbzPath..."
        try {
            Compress-Archive -Path (Join-Path -Path $tempDir -ChildPath "*") -DestinationPath $mergedCbzPath -Force
        }
        catch {
            Write-Error "Erreur lors de la compression en $mergedCbzPath : $_"
        }

        # Supprimer le répertoire temporaire de fusion
        Write-Host "Nettoyage des fichiers temporaires..."
        try {
            Remove-Item -Path $tempDir -Recurse -Force
        }
        catch {
            Write-Error "Impossible de supprimer le répertoire temporaire de fusion : $tempDir. Erreur : $_"
        }

        Write-Host "Volume $volumeNumber du Titre '$titre' fusionné avec succès en $mergedCbzPath`n"
    }

    Write-Host "Tous les volumes du manga '$($mangaDir.Name)' ont été traités.`n"
}

Write-Host "Tous les mangas ont été traités."
