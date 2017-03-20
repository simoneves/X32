# X32

This project grew out of a dissatisfaction with the on-board programming workflow of the X32/M32. I wanted a way to create cues for controlling the assignment of channels to DCAs, with labeling of the DCAs, and automatic muting and un-muting of channels, without touching *any* other parameter on the board.

The workflow presented here solves that problem, although in a rather technical way.

The user creates the DCA assignments in a spreadsheet. I used OpenOffice on a Mac, but I have had people successfully modify this script to read other formats of spreadsheet on other operating systems.

The Python script then reads the spreadsheet file and spits out a set of X32/M32 snippet files which can be transferred to the board on a USB stick. You set the board to snippet automation mode, turn off the prompt pref for scene recall, import the set of snippets starting from index 0, and then just hit the [GO] button.

Channels assigned to a DCA on any given cue (snippet) will be un-muted. Channels not assigned to a DCA will be muted. The DCA faders scribble strips will light up with the arbitrary text entered in the spreadsheet.

Other features allow the DCA control of bus masters and aux inputs, and (more experimentally) the control of bus send mutes for a particular channel (e.g. to enable reverb on a given vocal channel in a particular cue).

You will need a working Python installation on your computer, along with the *pyexcel_ods* module. I used Homebrew on my Mac to install a separate modifiable Python, then installed the required module with *pip*. The script should work fine on Windows and Linux too, but I'll have to leave it to you to get it running.

The files here are the Python script itself, a Bash script to run it with a certain set of parameters, an example spreadsheet (from the production of COMPANY that I'm opening this week) and example output snippet files.

The spreadsheet has to be in a certain format, and you have to edit the Python script to tell it which parts of the spreadsheet to look at.

SHEET_NAME
This is the default name of an OpenOffice spreadsheet document, as displayed in the tabs at the bottom. Change this string if you've changed the name

SKIP_ROWS
How many rows from the top to skip before attempting to interpret the spreadsheet contents

SNIPPET_COL / CUE_COL
The column numbers (starting from zero!) of the columns from which to pull the snippet number and the label string

FIRST_CHAN / NUM_CHANS
The first console channel to control (usually 1), and the number of channels to control (inclusive)

FIRST_CHAN_COL
The column number (starting from zero!) of the first console channel's data in the spreadsheet

(same for buses and aux-ins)

FIRST_DCA_COL
The column number (starting from zero!) of the first DCA's data in the spreadsheet (it is assumed starting from DCA 1)

NUM_DCAS
The number of DCAs to generate data for

I will describe the rest of the parameters in a later update, but the defaults should be fine for anyone else.

Download these files, make the ./runit.sh script executable (this script will only work on Mac or Linux), and run it.
