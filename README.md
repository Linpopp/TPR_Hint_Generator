# TPR_Hint_Generator
EN

The online version if you want : https://tprhint.thoughtless.eu

How to use :

1 - Install Python 3 on your PC (Windows store for example)

2 - Install the openpyxl, glob2 and jinja2 libraries. To do this, enter the following commands one by one in a CMD:
pip install openpyxl
pip install glob2
pip install jinja2

3 - Place the spoiler_log of the seed in the same folder as the script (be careful, there should only be one)

3 - Open a CMD at the root of the project (where the hint.py script is located)

4 - Type the following command: python3 hint.py

5 - The hint_list.txt and hint_list.xlsx files will be generated in the same folder as the script (do not run the script if one of these files is open on your PC)

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

1 - Defining checks in always

2 - Defining the dictionary that will contain the items of the always checks

3 - The script goes through the placement of items in the spoiler log to find the always checks and their item to store

4 - It displays the items contained by the checks

Sometimes:

1 - Defining checks in Sometimes

2 - Defining the dictionary that will contain the items of the Sometimes checks

3 - The script goes through the placement of items in the spoiler log to find the Sometimes checks and their item to store

4 - Selects 3 Sometimes randomly.

----------------------------------------------------------------------------------------------------------------------------------------------------------------
FR

La version en ligne si vous voulez : https://tprhint.thoughtless.eu

Utilisation :

1 - Installez Python 3 sur votre PC (Windows store par exemple)

2 - Installer les bibliothèques openpyxl, glob2 et jinja2. Pour cela, faites les commandes suivantes une à une dans un CMD :

pip install openpyxl
pip install glob2
pip install jinja2

3 - Placez le spoiler_log de la seed dans le même dossier que le script (attention, il ne doit y en avoir qu'un seul)

3 - Ouvrez un CMD à la racine du projet (là où se trouve le script hint.py)

4 - Tapez la commande suivante : python3 hint.py

5 - Les fichiers hint_list.txt et hint_list.xlsx seront générés dans le même dossier que le script (ne pas exécuter le script si un de ses fichiers est ouvert sur         votre PC

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

2 - Définition du dictionnaire qui contiendra les items des checks always

3 - Le script parcourt le placement des items du spoiler log pour trouver les checks always et leur item pour les stocker

4 - Il affiche les items contenu par les checks

Sometimes :

1 - Définition des checks en Sometimes

2 - Définition du dictionnaire qui contiendra les items des checks Sometimes

3 - Le script parcourt le placement des items du spoiler log pour trouver les checks Sometimes et leur item pour les stocker

4 - Sélectionne au hasard 3 Sometimes
