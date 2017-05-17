# X32Snippets

Current Version 1.5 Beta 3

### Background

This project grew out of a dissatisfaction with the on-board programming workflow of the X32/M32. I wanted a way to create cues for controlling the assignment of channels to DCAs, with labeling of the DCAs, and automatic muting and un-muting of channels, without touching *any* other parameter on the board.

The workflow presented here solves that problem, although in a rather technical way.

The user creates the DCA assignments in a spreadsheet. I used OpenOffice on a Mac, but I have had people successfully modify this script to read other formats of spreadsheet on other operating systems.

The Python script then reads the spreadsheet file and spits out a set of X32/M32 "Snippet" files, plus a parent "Show" file, which can be transferred to the board on a USB stick. You set the board to "Cues" automation mode, turn off the prompt pref for scene recall, import the "show" file, and then just hit the `GO` button.

Channels assigned to a DCA on any given cue (snippet) will be un-muted. Channels not assigned to a DCA will be muted. The DCA faders scribble strips will light up with the arbitrary text entered in the spreadsheet.

Other features allow the DCA control of bus masters and aux inputs.

Additional custom functionality requested by users is described later in this document.

You will need a working Python installation on your computer, along with the *pyexcel_ods* module. I used Homebrew on my Mac to install a separate modifiable Python, then installed the required module with *pip*. The script should work fine on Windows and Linux too, but I'll have to leave it to you to get it running.

The files here are the Python script itself, a Bash script to run it with a certain set of parameters, an example spreadsheet (from the production of COMPANY that I'm opening this week) and example output snippet files.

The spreadsheet has to be in a certain format, and you have to edit the Python script to tell it which parts of the spreadsheet to look at.

### Control Parameters

**SHEET_NAME**
This is the default name of an OpenOffice spreadsheet document, as displayed in the tabs at the bottom. Change this string if you've changed the name

**SKIP_ROWS**
How many rows from the top to skip before attempting to interpret the spreadsheet contents

**CUE_NUM_COL / CUE_LABEL_COL**
The column numbers (starting from zero!) of the columns from which to pull the cue number and the cue label string

**FIRST_CHAN / NUM_CHANS**
The first console channel to control (usually 1), and the number of channels to control (inclusive, must be 32 or less)

**FIRST_CHAN_COL**
The column number (starting from zero!) of the first console channel's data in the spreadsheet

(same for buses and aux-ins) (NUM must be 16 or less, or 6 or less, respectively)

**FIRST_DCA_COL**
The column number (starting from zero!) of the first DCA's data in the spreadsheet (it is assumed starting from DCA 1)

**NUM_DCAS**
The number of DCAs to generate data for (must be 8 or less)

**DCA_COLOR**
The desired backlight color for the DCA labels (when active) (can be `RD`, `GN`, `YE`, `BL`, `MG`, `CY`, `WH`, `RDi`, `GNi`, `YEi`, `BLi`, `MGi`, `CYi`, `WHi`)

### Usage

Download these files, edit the ./runit.sh script appropriately, make it executable (this script will only work on Mac or Linux), and run it.

NOTE: The resulting files have only been tested to load successfully on a console running firmware v3.07.

### Craig Flint functionality

Enabling **CF_WARN_DCAS** will cause DCA faders which will be active on the *next* cue to be illuminated with the color specified in **CF_WARN_COLOR** (value options same as above). Note that this code only checks the next spreadsheet row, so will not detect the next active DCA if there are row gaps between cues.

Enabling **CF_NAME_CHANS** will cause the channel labels to be updated as well as the DCA labels. This requires an additional set of columns in the spreadsheet, starting at **CF_FIRST_CHAN_NAME_COL**. If a cell in those columns is non-empty, the channel name will be changed to that value on that cue.

Enabling **CF_ALT_LABEL_COLORS** will cause DCA labels which match any of the strings in **CF_ALT_LABELS** to be colored as **CF_ALT_COLOR** instead of **DCA_COLOR**.

### STC functionality

Enabling **STC_FX** will additionally control the bus send mute for any given channel to the single FX bus specified by **STC_FX_BUS**. If the DCA number in the channel column is negative, the channel will still be assigned to the DCA of the positive value, but the bus send for that channel will be un-muted. This allows, for example, reverb to be selectively applied to channels in a given cue, and not in others.
