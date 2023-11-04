## Fonctionnalité du Script

Le script permet d'insérer des images dans un fichier Excel en liant le contenu d'une cellule à un nom de fichier JPG stocké dans un dossier spécifié.

Le script suit les étapes suivantes :

1. Désactive la mise à jour de l'écran pour améliorer les performances.
2. Gère les erreurs en redirigeant vers l'étiquette "ErrorHandler" pour la gestion des erreurs.
3. Déclare les variables nécessaires.
4. Définit le dossier où sont stockées les images et le type de fichier (extension).
5. Définit la plage de données dans la colonne A de la deuxième ligne jusqu'à la dernière cellule non vide.
6. Initialise une variable pour numéroter les images.
7. Supprime les images existantes dans la feuille de calcul.
8. Parcourt chaque cellule de la plage définie.
9. Pour chaque cellule non vide, il récupère la valeur de la cellule, construit le chemin complet du fichier image en utilisant le dossier, le nom de fichier et l'extension, puis vérifie si le fichier image existe.
10. Si le fichier image existe, il crée une référence à la cellule de destination dans la colonne B.
11. Ensuite, il insère l'image dans le fichier Excel en utilisant la méthode "AddPicture" et configure les propriétés de l'image, telles que la position, la taille, le nom, le verrouillage du rapport d'aspect, etc.
12. Redimensionne la cellule de destination pour s'adapter à l'image.
13. Déplace l'image avec les cellules pour qu'elle reste alignée lors du déplacement des cellules.
14. Centre le texte au milieu de la cellule.
15. Incrémente le numéro d'image.
16. Après avoir traité toutes les cellules, il réactive la mise à jour de l'écran et se termine.
17. En cas d'erreur, il affiche un message d'erreur avec une description de l'erreur et réactive la mise à jour de l'écran.

Ce script est utile lorsque vous avez un fichier Excel contenant une liste de noms de fichiers d'images, et vous souhaitez les afficher dans le fichier Excel en les liant aux valeurs des cellules.
