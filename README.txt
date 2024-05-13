------------------------------------------
BOSSMAN JACK SIMULATOR IN EXCEL 2023 demo
First Entry: 4/21/2024
Updated 1  : 4/28/2024
------------------------------------------

The basic idea of this repository is to mimic few popular or custom gambling games and code the hypothethical reaction to them of ex-kick streamer Austin Peterson, aka Bossman Jack. Project is written in VBA (using Microsoft Excel 2021 Software), and assets are extracted from various youtuber's clips and videos of Bossman Jack (modified via Python opencv2 and moviepy libraries). 
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
There is stored there simple analyssis of reactions, each reaction source with author and timestamps from videos.

As for today, there are implemented following games:
	1) YesNo     - it is basically a coin flip, but with 25 % chance of winning, based on this reddit post:
		       https://www.reddit.com/r/bossmanjack/comments/1bqfwqo/new_game_made_just_for_the_boss/,
	2) Mines     - minesweeper, but a little bit more rigged.
	3) Keno      - random number generator on steroids - from 40 cells player must hit the highest ammount possible from 
		       10 chosen. More chosen cells hitted - the bigger payout.

-----------
Project is in EARLY version.
The project is made clearly for humorous purposes, it is goal is not to spread misinformation or any malicious statements about said streamer.
Any hostile related to Austin thing this project is supposed to do, is to warn people against the dangers of gambling addiction. 
Please, do not also use it for harassing, doxing and performing any illegal action towards the Bosmman.

-----------
HOW TO RUN?
In order to run this project, you need a Windows OS with winmm.dll library (it is used for reading the music files). Also you need to have
Microsoft Excel with enabled macros (without it, it won't work). Download the content of repository, save it somewhere on your drive, unpack 
assets file in the same directory where there is stored .xlsm file, and you are ready to run.

-----------
OTHER INFO
If you have any ideas how to improve the project or got any more clips you wanted to be included in project, please contact with me on my reddit page https://www.reddit.com/user/RotianQaNWX/.