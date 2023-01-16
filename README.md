# TPR_Hint_Generator

How to use :


Script explanation:

WotH:
1 - Define a list of important items
2 - The script loops through the spheres in the spoiler log, and for each match with the list and the item names in them, it puts the related checks for that item into an "important" group.
3 - Create a dictionary that stores for each game check the "zone it is associated with" (I based this on the gdoc)
4 - For each check in the important zones list, the script will look at the zone that corresponds to that check and bring it up in a list
5 - It selects 3 randomly.

Barren:
1 - The script stores all the zone names in a list
2 - It creates the variable that will store the barren zones
3 - It puts all the content of the zone list in this variable, excluding the WotH zones
4 - It creates a variable that lists all the dungeon names
5 - It excludes from this variable the dungeons required to have the list of barren dungeons
6 - It excludes from the Barren zone variable the content of the barren dungeon variable
7 - It selects 2 randomly.

Always:
1 - Definition of checks as always
2 - Definition of the list that will contain the items of the always checks
3 - The script loops through the placement of the items in the spoiler log to find the always checks and their item
4 - It displays the items contained by the checks

Sometimes:
1 - Creation of a dictionary that stores all the names of dungeon checks and the name of the dungeon they are associated with
2 - Declaration of a variable that will store all the dungeon checks that are barren
3 - The script will look in the variable that contains the names of the barren dungeons (created previously in step 4 and 5 of "Barren") for the names of the corresponding checks for those dungeons.
4 - It will include them in the variable from step 2
5 - Creation of a list that stores all the checks excluded in the ruleset
6 - Selects 3 checks randomly from the item placement list in the spoiler log, excluding the barren dungeon checks and those excluded in the ruleset.

----------------------------------------------------------------------------------------------------------------------------------------------------------------

Explication du script :

WotH :
1 - Définition une liste des items importants
2 - Le script parcourt les spheres du spoiler log, et pour chaque correspondance avec la liste et les
    noms d'items dans celles-ci, il met les checks liés à cet item dans un groupe "important".
3 - Création d'un dictionnaire qui stock pour chaque check du jeu la "zone qui lui est rattachée",
    (je me suis basé sur le gdoc pour ça)
4 - Pour chaque check de la liste des zones importantes, il script va regarder la zone qui correspond
    à ce check et la remonter dans une liste
5 - Il en sélectionner 3 au hasard.

Barren :
1 - Le script stock tout les noms de zones dans une liste
2 - Il créer la variable qui va stocker les zones barren
3 - Il met dans cette variable tous le contenu de la liste des zones, j'exclus de cette variable les
    zones WotH
4 - Il créer une variable qui liste tous les noms de donjons
5 - Il exclus de cette variable les donjons requis pour avoir la liste des donjons barren
6 - Il exclus de la variable des zones Barren le contenu de la variable des donjons barren
7 - Il en sélectionne 2 au hasard

Always :
1 - Définition des checks en always
2 - Définition de la liste qui contiendra les items des checks always
3 - Le script parcourt le placement des items du spoiler log pour trouver les checks always et leur item
4 - Il affiche les items contenu par les checks

Sometimes :
1 - Création d'un dictionnaire qui stock tous les noms de checks de donjons et le nom du donjon qui
    leur est rattaché
2 - Déclaration d'une variable qui va stocker tous les checks des donjons qui sont barren
3 - Le script va chercher dans la variable qui contient les noms des donjons barren (créée précédemment
    dans l'étape 4 et 5 de "Barren") les noms des checks correspondants à ces donjons.
4 - Il va les inclure dans la variable de l'étape 2
5 - Création d'une liste qui stock tous les checks exclus dans le ruleset
6 - Sélectionne au hasard 3 checks de la liste du placement des items du spoiler log, en excluant les
    checks des donjons barren et ceux exclus dans le ruleset.
