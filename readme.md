﻿# The Last Crusade
I tried to search Github for this game, but I couldn't find it anywhere. As such, I decided to host it here.
## What is it?
The Last Crusade is a game made by Patrick Dwyer & Peter S. VanLund. The game was made as an example of what the game engine and map maker they created could do. Unfortunately, the map maker wasn't used very much because it was not accessible to screen readers, so beyond the game itself the engine and map maker never gained much traction. This may have also been due to the lack of documentation on the engine, and as such I unfortunately cannot tell you what dependencies are required to work with the source code or even what license the code is under. Feel free to update this repo if you guys know anything I don't. I will, however, provide the descriptions they gave on their website in the hopes that it might help the more code savvy part of the community figure out how to make changes. The page for their RPG Engine can be found [here](http://www.cs.unc.edu/Research/assist/et/projects/RPG/index.html). Pull requests are very much welcome. When making a pull request, please provide a build version containing the same changes you made in the source. Enjoy.
# RPG Game Engine and Map Maker
Patrick Dwyer & Peter S. VanLund
## Purpose:
To create a role-playing game without any visuals.
## Reason:
We grew up playing RPG's. There should be a Gaming Engine that reads out the events of the game instead of using text or visuals to display the actions.  In order to do this we created a Game Engine and a Map Maker.
## Gaming Engine:
We based our RPG game engine off of text-based Role-Playing Games.  However, our engine does not print to the screen the actions that are taking place.  It simply reads out the events as they occur.  Therefore, people that have the ability to hear can easily play the game.  The Engine is basically a huge state machine.  It reads in the user input and processes their action and acts accordingly.  However, our Game Engine only works on maps created in our Map Maker.
## Map Maker:
Our Map Maker helps create maps that the Game Engine can read in and set up for the user to play.  The Map Maker allows for a user to create different weapons, armor, items, creatures, friends, vendors, and even leprechauns.  Unfortunately, our Map Maker is not accessible for the visually impaired.  However, the Map Maker part of our project ensures that anyone who creates a map can easily play it using our Game Engine.  This fact makes are game easily extensible.  You can create an unlimited number of levels to play with characters and items that you name and give stats to.  You can even put in your own music.
## Game Play:
- The user is a character that has strength, defense, hit points, gold, potions, and an item list.
- There are five types of items: weapons, armor, gold, potion, and special items.  Weapons will increase your strength.  Strength is used to determine how much you attack an enemy for.  Armor will increase your defense.  Defense is the amount you defend for in combat.  Gold is used to purchase items from vendors.  Your hit points are your health or life.  When you have zero hit points, you die.  Using a potion will add some health to you after you have been attacked.  The item list holds special items.  In the Map Maker, you can set certain locations to be locked unless a user has a certain item.  A special item could be a "key" that the user finds.  All items have a name sound and an action sound.  The name sound is read when you encounter that item.  The action sound is played when you use that item.
- There are four types of characters: enemies, friends, vendors, and leprechauns.  Enemies have strength, defense, hit points, runaway percentage, and an item list.  You interact with these characters by fighting or running away from them.  If you defeat an enemy, you will get the items that are in the enemy’s item list.  Friends are characters that will give you advice or items if you choose to talk to them.  Vendors are characters that will sell you items for gold.  Their strength value is used to determine how good their prices are for weapons and armor.  Their defense value is used to determine how good their prices are for potions.  Their item list is used to show what items they have for sale.  Leprechauns are special characters that you play a gambling game with to win gold.  All characters have a name sound and an action sound.  The name sound is read when you encounter that character.  The action sound is different for each character.
- Interaction with an enemy.  When you encounter an enemy, the game will ask you if you choose to run away.  If you succeed you will return to the previous location.  Combat is turn based.  The enemy will attack and then you will have the chance to respond.  If you fail to run away, the enemy will attack you.  To figure out how much an attack hit for, the Game Engine gets a random number from the first character's strength to two times that character's strength.  You then subtract the second character's defense from that random number.  The action sound for the enemy is the sound that is played when the enemy attacks you.  Your attack will be the action sound of the weapon that you currently have equipped.
- Interaction with a friend.  When you encounter a friend, the game will ask you if you would like to talk to them.  You respond yes or no.  If yes, then the friend will read out some advice.  The advice for a friend is his action sound.  If no, then nothing happens.
- Interaction with a vendor.  When you encounter a vendor, the game will ask you if you would like to purchase some items.  If yes, then the game will read out his list of items for sale and the price of those items one at a time.  At the end of each item and price, the game will ask you if you would like to purchase that item.
- Interaction with a leprechaun.  When you encounter a leprechaun, the game will ask you if you would like to play a game with him.  If yes, then you will pick a number one, two, or three.  Sometimes you win some gold, and sometimes you lose.  Leprechauns will sometimes get angry if you do not play the game.
- Interaction with locations.  You can do two things at a location.  You can search a location for items, and you can find out which directions are open or have been visited.
- Saving the game.  There can be only one saved game.  Therefore, when you save the game, you are overwriting your previously saved game.
## Keyboard Input:
- spacebar: reads out your stats and the instructions for playing
- arrow keys up, down, right, and left: respectively move North, South, East, and West
- a: will attack an enemy
- r: will try to run away from an enemy
- s: will search a location
- d: will tell you which directions are open or which directions have been visited
- p: will use a potion
- f: will skip over speeches and dialogues
- y: signifies yes
- n: signifies no
- v: will save the game
- 1, 2, and 3: respectively signify 1, 2, and 3
## Using the Map Maker:
- To create a map, open a new map and name it whatever you want.  The nodes are locations.  The first node you click on will be the starting node for the map.  Every node that you click on, will be another location that the player can go into.
- To add characters to the map, you must first add characters to the Map Maker.  Click on the add button and type in a character's name that you would like to put in.  You then select his character type.  You can then choose the stats for that character.  The stats are based on ranges.  For example, 5 to 7 strength will randomly pick 5,6, or 7 for that character's strength when the map loads.  This makes the game random when you play. 
- To add the character to a location, put the cursor on the location that you would like to add him to. You then go to Node Properties and click on your character.  You can then set the percentage or probability of that character being at that node.  There will only be one character for each node.  Therefore, if you want there to be a random chance of three different characters, make sure that the percentages of those characters do not total over 100. 
- To add a name sound to the character, click on the add sound button first.  Then, search for the filename of your name sound. After the sound has been added, you can then go to character properties and add the name sound. The action sound works the same way. 
- To add an item to the map, click the add item button and type in a name.  You then select the item type and value.  You can add items to characters by clicking on a character and setting the probability of that character having this item.  You do not have to worry about percentages totaling over 100 here because characters can have multiple items.  You can add items to locations in the same way that you added characters to the locations.  Once again you do not have to worry about the percentages totaling over 100 because locations can have more than one item.  Just use the Node properties section of the Map Maker.  Add sounds to the items in the same way that you added sounds to the characters. 
- To add music to a location, you must first add the music to the Map Maker.  Then, in the Node properties part of the Map Maker, click on the music that you want for that location.  You can click on multiple music files, the game will randomly select one when the player arrives at that node. 
- To set a required item for a location, first bring the cursor to the desired node.  Then check the item that you want to be required.  The player cannot enter that node unless they have the required item.
- When you are done with your map, make sure that you have an end node.  To do this, put the cursor on the node that you want, and click the make end node button.  You can then save your map. 
- To play your newly created map, select Save and Play in the Map Maker.  The only difference in gameplay is that you cannot save your game.
## Our game:
Using the Game Engine and the Map Maker, we created our very own RPG.
# The Last Crusade
"In a time of great despair, King Vaisara called for the finest warriors of the land. He sent them on a dangerous journey. Their mission was to kill the evil King Lasarus. None of those warriors have returned. It is now, in the darkest hour of the Vaisara Kingdom, that King Lasarus must be stopped. King Vaisara has requested that you bring peace back to this kingdom. Your journey will begin in the village of Cascata. Travel through the graveyard and reach Castle Lasarus. There you must kill King Lasarus. The fate of the Vaisara Kingdom is in your hands."
## Download
- [Game.zip](http://www.cs.unc.edu/Research/assist/et/projects/RPG/Game.zip) (76.8 MB): Download to play game.  Includes maps, sounds, and game & map editor executables (+fmod.dll).
- [Source.zip](http://www.cs.unc.edu/Research/assist/et/projects/RPG/Source.zip) (249 KB) : Includes source code for both game engine & map editor, including FMOD headers, .lib, & .dll.
- [VB Runtime files](http://www.cs.unc.edu/Research/assist/et/projects/RPG/VB.exe) (1.18 MB): Includes Visual Basic runtime files needed to use Map Maker.
### (Note from brogar2000
All 3 resources from the above links are included in this Github package.)
## Minimum System Requirements
- CPU:   500 MHz
- RAM:  256 MB
- Audio: sound card with speakers

## The World:
- Universal Items
 - Gold coins 50
 - Bag of gold 100
 - Potion 25
- Level 1: The Village of Cascata
 - Enemies:
  - Wolf
  - Berserker
  - Snake
  - Bear
  - Barbarian
  - Centaur
  - Giant Spider
  - Cyclops
 - Friends
  - Fisherman
  - Farmer
 - Vendors
  - Nice Price Weapons
  - Weapons R Us
 - Leprechauns: 2 inhabit Cascata
 - Weapons
  - Wooden Dagger 2
  - Small Dagger 3
  - Dagger 4
  - Sharp Dagger 5
  - Silver Dagger 7
  - Black Talon 8
  - Small Sword 10
 - Armor
  - Wooden Shield 2
  - Small Shield 3
  - Shield 5
  - Strong Shield 7
  - Weak Armor 10
 - Gold
  - Pot of Gold 250
- Level 2: The Graveyard
 - Enemies
  - Bat
  - Skeleton
  - Zombie
  - Werewolf
  - Vampire
  - Nightmare
  - Death Knight
  - Necromancer
 - Friends
  - Priest
  - Spirit
 - Vendors
  - Deadly Discounts
  - Death Dealers
 - Weapons
  - Small Sword 10
  - Sword 13
  - Long Sword 15
  - Death Blade 18
  - Small Axe 20
 - Armor
  - Weak Armor 10
  - Light Armor 12
  - Basic Armor 14
  - Medium Armor 16
  - Heavy Armor 18
  - Weak Helmet 20
 - Special Items
  - Skeleton Key
  - Undead Elixir
- Level 3: The Castle
 - Enemies
  - Guard
  - Swordsman
  - Griffin
  - Silver Knight
  - Prince Lasarus
  - King Lasarus
 - Friends
  - The Golden Knight
 - Weapons
  - Spear 30
  - Double Edged Blade 35
  - Knight's Sword 40
  - Royal Blade 50
 - Armor
  - Guard's Armor 45
  - Swordsman's Armor 50
  - Knight's Armor 60
  - Royal Armor 75
 - Special Items
  - King Lasarus's Crown
