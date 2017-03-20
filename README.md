# X32

This project grew out of a dissatisfaction with the on-board programming workflow of the X32/M32. I wanted a way to create cues for controlling the assignment of channels to DCAs, with labeling of the DCAs, and automatic muting and un-muting of channels, without touching *any* other parameter on the board.

The workflow presented here solves that problem, although in a rather technical way.

The user creates the DCA assignments in a spreadsheet. I used OpenOffice on a Mac, but I have had people successfully modify this script to read other formats of spreadsheet on other operating systems.

The Python script then reads the spreadsheet file and spits out a set of X32/M32 snippet files which can be transferred to the board on a USB stick. You set the board to snippet automation mode, turn off the prompt pref for scene recall, import the set of snippets starting from index 0, and then just hit the [GO] button.

Channels assigned to a DCA on any given cue (snippet) will be un-muted. Channels not assigned to a DCA will be muted. The DCA faders scribble strips will light up with the arbitrary text entered in the spreadsheet.

Other features allow the DCA control of bus masters and aux inputs, and (more experimentally) the control of bus send mutes for a particular channel (e.g. to enable reverb on a given vocal channel in a particular cue).

You will need a working Python installation on your computer, along with the *pyexcel_ods* module. I used Homebrew on my Mac to install a separate modifiable Python, then installed the required module with *pip*. The script should work fine on Windows and Linux too, but I'll have to leave it to you to get it running.

python2.7 ./X32Snippets.py spreadsheet.ods SHOWNAME
