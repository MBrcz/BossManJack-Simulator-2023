------------------------------------------
BOSSMAN JACK SIMULATOR IN EXCEL 2023 demo
First Entry: 4/21/2024
Updated 1  : 4/28/2024
------------------------------------------

The basic idea of this repository is to mimic few popular or custom gambling games and code the hypothethical reaction to them
of ex-kick streamer Austin Peterson, aka Bossman Jack. Project is written in VBA (using Microsoft Excel 2021 Software), and assets are
extracted from various youtuber's clips and videos of Bossman Jack (modified via Python opencv2 and moviepy libraries). 
Youtubers, which clips I used for this project:
	a) https://www.youtube.com/@ColonelRat,
	b) https://www.youtube.com/@Austin_07Schadenfreude,
	c) https://www.youtube.com/@LOSSmanJack,
	d) many others...

There are for types reactions of Bossman can have for shown events:
	a) win [currently 7 different clips],
	b) loss [currently 5 different clips],
	c) beg [currently 4 different clips],
	d) rage [currently 6 different clips],
	e) bigwin [currently 1 different clips, NOT IMPLEMENTED YET].

If you are intrested more about the clips and the source material - check the ~/Reactions Table.xlsx file.

As for today, there are implemented following games:
	1) YesNo     - it is basically a coin flip, but with 25 % chance of winning, based on this reddit post:
		       https://www.reddit.com/r/bossmanjack/comments/1bqfwqo/new_game_made_just_for_the_boss/,

	2) Mines     - it works similary to minesweeper, but with this difference, that in case of not finding mine, player
		       do not get visual feedback when mine can be found. The more not mines (aka Gems) player can find, the more money will he get by         
                       cashing out (multiplier is 1.05 of waged money for each gem found). In case of finding a mine, game is ending and player loses all 		       		       waged money. Also it is worth noting, that this game is agressively rigged against the player (this is author design choice).
		       
		       In normal case, durning the initialization of the game, the game should be storing a 2D array with positions of mines and gems.
                       However, in this case, game decides whether player has found a mine ad hoc using a formula: RND() * 100 >= 30 + (Gems found * 10), where
                       RND() is a random number from 0 to 1. Therefore it is not likely player will ever score more than 7 gems. In case of		               		       	       encountering the mine, game ad hoc invents "position" that player was moving around current session. Therefore, it does not matter where player			       chooses cells, and how many mines are present on the board. The cashout is calculated as 
                       = (waged money mult x (0.1 * found gems)) * found gems * waged money 
		       Waged money mult is a constans and it's equal to 1.05.
	
	3) Keno      - I do not know who, invented this game, but this person should be locked deep inside a prison and isolated from society. It is HIGHLY  		       		       additcive (even for the standards of traditional gambling games). The game consists of board with 40 cells numered 1 to 40. Durning the 				       initialization of the board 10 cells are randomly chosen to be a "target cells" (they are in violet/purple color). Through selecting play button, 		       player draws 10 random cells from the board, the more target cells shall be drawn, the bigger payback will. For the drawing procedure is 			       responsible software (code).
		       The table for payback in Keno in this case are as follows:

                       Multiplier || 0.00x  ||  0.00x  ||  0.00x  ||  0.00x  ||  3.50x  ||  8.00x  ||  13.00x  ||  63.00x  ||  500.00x  ||  800.00x  || 1000.00x
           	       ----------------------------------------------------------------------------------------------------------------------------------------
		       Target hit || 0x     ||  1x     ||  2x     ||  3x     ||  4x     ||  5x     ||  6x      ||  7x      ||  8x       ||  9x       || 10x		       
		       
		       The ammount of won money is calculated as (multiplier * waged money). In order to skew game a little (becouse durning playtests I noticed
		       that wins (3.5x+) are happening definetelly too often and too consistently [or I and my tester were especially lucky]), whenever the game founds a 		       target cell, there is 20% chance that it will be discarded and draw will be repeated for this current iteration.

	It is worth noting, that those games are only my own implentation, real casinos might use much more complicated and dirtier tricks in order to skew games (like 	diminishing returns on wins, causing fake win streaks and big payouts for newer accounts, etc, manipulating accounts chances on getting win / losses.

-----------
Project is in Alpha version.
The project is made clearly for humorous purposes, it is goal is not to spread misinformation or any malicious statements about said streamer.
Any hostile related to Austin thing this project is supposed to do, is to warn people against the dangers of gambling addiction. Please, do not also use it for harassing, doxing and performing any illegal action towards the Bosmman.

-----------
HOW TO RUN?
In order to run this project, you need a Windows OS with winmm.dll library (it is used for reading the music files). Also you need to have
Microsoft Excel with enabled macros (without it, it won't work). Download the content of repository, save it somewhere on your drive, unpack 
assets file in the same directory where there is stored .xlsm file, and you are ready to run.

-----------
OTHER INFO
If you have any ideas how to improve the project or got any more clips you wanted to be included in project, please contact with me on my reddit page https://www.reddit.com/user/RotianQaNWX/.