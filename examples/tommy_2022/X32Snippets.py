#!/usr/bin/env python

################################################################################
#
# X32 Snippets
#
# Last mod:
# January 22, 2022
#
# Written by:
# Simon Eves (simon@eves.us)
#
# Additional contributions by:
# Craig Flint (craig@cliveflint.co.uk)
# Aaron Yoffe (aaron@blownfuse.org)
# Art Zemon (art@zemon.name)
#
# Free for non-commercial use
#
################################################################################

VERSION = "2.0" # port to Python 3

################################################################################
# Imports
################################################################################

import time
import sys
import string

from pyexcel_ods import get_data


################################################################################
# Constants
# Set these appropriately for your source spreadsheet layout
################################################################################

SHEET_NAME                 = 'Sheet1' # spreadsheet sub-sheet name containing all data

SKIP_ROWS                  = 4     # number of spreadsheet rows to skip before extracting data

CUE_NUM_COL                = 2     # spreadsheet column of cue number data (can be integer or tenths, e.g. "1" or "1.1")
CUE_LABEL_COL              = 3     # spreadsheet column of cue data (any short string)

PATH_NUM_ROW               = 1     # row containing path (chan, bus, aux) numbers

FIRST_CHAN_COL             = 15    # spreadsheet column of data for that first channel
NUM_CHANS                  = 19    # number of contiguous set of channels

FIRST_BUS_COL              = 0     # spreadsheet column of data for that first bus
NUM_BUSES                  = 0     # number of contiguous set of buses (4 = FX1-4)

FIRST_AUXIN_COL            = 0     # spreadsheet column of data for that first auxin
NUM_AUXINS                 = 0     # number of contiguous set of auxins (2 = AuxIn 5-6)

FIRST_DCA_COL              = 5     # spreadsheet column of data for first DCA
NUM_DCAS                   = 8     # number of DCAs
DCA_COLOR                  = 'WH'  # color for active DCA labels

NAME_CHANS                 = False # set channel names
FIRST_CHAN_NAME_COL        = 0     # spreadsheet column of name for the first channel

DCA_ALT_LABEL_COLORS       = False
DCA_ALT_LABELS             = [ 'Reverb', 'Sh Dly', 'Lg Dly', 'Delay' ] # color these DCA labels differently
DCA_ALT_LABEL_COLOR        = 'MG'

DCA_ACTIVE_ON_NEXT_CUE     = False # alternative color for DCAs that are still active on the next cue
DCA_ACTIVE_ON_NEXT_CUE_COLOR = 'RD'  # that color

DCA_SAME_ON_NEXT_CUE       = False # alternative color for DCAs that stay the same on the next cue
DCA_SAME_ON_NEXT_CUE_COLOR = 'GN'  # that color

FX_UNMUTE                  = True  # negative path DCA indexes mean also un-mute that path's FX send to the bus below
FX_UNMUTE_BUS              = 15    # set the send from the paths to this single FX bus

OTHER_MUTES                = False # mute one or more band channel per cue, all others will be unmuted (e.g. for switching basses)
OTHER_MUTES_FIRST_CHAN     = 17    # first board channel for the band
OTHER_MUTES_NUM_CHANS      = 8     # number of consecutive board channels for the band
OTHER_MUTES_COLS           = [ ]   # spreadsheet column of data for first channel to mute

################################################################################
# Functions
################################################################################

def read_cell_as_string(d, r, c):

    # robust function to pull contents of a spreadsheet cell as a string, regardless of actual contents
    # incoming indices are 1-based

    try:
        #print("DEBUG: Accessing row " + str(r) + ", cell " + str(c))
        row = d[r - 1]
        cell = row[c - 1]
        #print("DEBUG: row " + str(r) + ", cell " + str(c) + " is type " + str(type(cell)))
        #if type(cell) == unicode:
        #    print("DEBUG:  unicode")
        #    cell = cell.encode('utf-8')
        #else:
        #    print("DEBUG:  not unicode")
        s = str(cell)
        #print("DEBUG: '" + s + "'")
        return s
    except:
        #print("DEBUG: Exception!")
        return ''

def string_to_int(s):
    
    return int(float(s))

def process_paths(ods, snp_file, row_index, first_path_col, num_paths, osc_prefix):

    # general function to process paths of any of the three types
    # we pull the actual board path number from PATH_NUM_ROW

    # first we find which paths in this cue have any DCA assignment and therefore need to be unmuted
    mutes = []
    for path in range(0, num_paths):
        dca = read_cell_as_string(ods, row_index, path + first_path_col)
        if dca != '':
            mutes.append(False)
        else:
            mutes.append(True)
            
    # then we write out mute-ons for paths which have become muted
    for path in range(0, num_paths):
        if mutes[path]:
            path_num = string_to_int(read_cell_as_string(ods, PATH_NUM_ROW, path + first_path_col))
            snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/mix/on OFF\n')

    # then we write out the new DCA assignments
    for path in range(0, num_paths):
        dca = read_cell_as_string(ods, row_index, path + first_path_col)
        bitmap = 0
        if dca != '':
            bitmap = 1 << (abs(string_to_int(dca)) - 1)
        else:
            bitmap = 0
        path_num = string_to_int(read_cell_as_string(ods, PATH_NUM_ROW, path + first_path_col))
        snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/grp/dca ' + str(bitmap) + '\n')
    
    # then we write out mute-offs for paths which have become un-muted
    for path in range(0, num_paths):
        if not mutes[path]:
            path_num = string_to_int(read_cell_as_string(ods, PATH_NUM_ROW, path + first_path_col))
            snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/mix/on ON\n')
            
def next_dca_label(ods, row_index, col):

    # function to find the contents of the cell in col 'col' in the next row below 'row_index' that has a valid cue number

    # are we already at the END?
    search_cue = read_cell_as_string(ods, row_index, CUE_NUM_COL)
    if search_cue == 'END':
        return ''

    # find the next row that has a valid cue
    search_row = row_index + 1
    search_cue = read_cell_as_string(ods, search_row, CUE_NUM_COL)
    while search_cue == '':
        search_row = search_row + 1
        search_cue = read_cell_as_string(ods, search_row, CUE_NUM_COL)

    # was it the END
    if search_cue == 'END':
        return ''

    # otherwise return whatever's in the column
    return read_cell_as_string(ods, search_row, col)

def current_or_previous_channel_name(ods, row_index, col):

    # function to find the contents of the next non-empty cell in col 'col' and row 'row_index' or above

    # are we already at the top?
    if row_index <= SKIP_ROWS:
        return ''

    # search upwards to find the previous non-empty cell in this column
    search_row = row_index
    search_name = read_cell_as_string(ods, search_row, col)
    while search_name == '' and search_row > SKIP_ROWS:
        search_row = search_row - 1
        search_name = read_cell_as_string(ods, search_row, col)

    # return
    return search_name


################################################################################
# Main
################################################################################    
        
if __name__ == "__main__":

    title = '# X32 Snippets v' + VERSION
    print('#' * len(title))
    print(title)
    print('#' * len(title))
    
    #
    # validate command-line parameters
    #
    
    if len(sys.argv) != 3:
        print("");
        print("Usage: X32Snippets.py <ods_file_name> <show_name>")
        sys.exit(0)
        
    #
    # get command line parameters
    #

    ods_file_name = sys.argv[1]
    show_name = sys.argv[2]

    #
    # process file
    #
    
    # report
    print("Opening spreadsheet...")
    
    # read the file
    ods = get_data(ods_file_name)[SHEET_NAME]
    
    # init these
    row_index = SKIP_ROWS
    snp_index = 0
    cue_numbers = []
    cue_labels = []
    
    # report
    print("Creating cues...")
    
    # iterate rows until the end
    for row_index in range(SKIP_ROWS + 1, len(ods)):
        
        # get cue
        cue = read_cell_as_string(ods, row_index, CUE_NUM_COL)
        
        # skip rows with no cue
        if cue == '':
            print("DEBUG: Skipping row " + str(row_index) + " with no cue")
            continue
            
        # the end?
        if cue == 'END':
            print("DEBUG: Found END, stopping")
            break
            
        # get cue number
        cue_number = ''
        try:
            # this could be more robust
            # also does not handle cues of form X.Y.Z or X.Y where Y > 9
            cue_number = str(int(round(float(cue) * 100.0)))
        except:
            print("ERROR: Found invalid cue number at row " + str(row_index + 1))
            sys.exit()
            
        # get cue label
        cue_label = read_cell_as_string(ods, row_index, CUE_LABEL_COL)
        
        # report snippet
        print('Generating new cue "' + cue + '", label "' + cue_label + '"')
            
        # open snippet file
        snp_file = open(show_name + '.' + str(snp_index).zfill(3) + '.snp', 'w')
        
        # start snippet
        snp_file.write('#2.1# "' + cue + '" 0 0 0 0 0\n')
        
        # process channels
        if NUM_CHANS > 0:
            process_paths(ods, snp_file, row_index, FIRST_CHAN_COL, NUM_CHANS, 'ch')
        
        # process buses
        if NUM_BUSES > 0:
            process_paths(ods, snp_file, row_index, FIRST_BUS_COL, NUM_BUSES, 'bus')
        
        # process auxins
        if NUM_AUXINS > 0:
            process_paths(ods, snp_file, row_index, FIRST_AUXIN_COL, NUM_AUXINS, 'auxin')
        
        # for channels only, also control the mute of the given FX bus send for this path
        if FX_UNMUTE:
            for chan in range(0, NUM_CHANS):
                dca = read_cell_as_string(ods, row_index, chan + FIRST_CHAN_COL)
                fx_on = False
                if dca != '':
                    fx_on = string_to_int(dca) < 0
                chan_num = string_to_int(read_cell_as_string(ods, PATH_NUM_ROW, chan + FIRST_CHAN_COL))
                if fx_on:
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/mix/' + str(FX_UNMUTE_BUS).zfill(2) + ' ON\n')
                else:
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/mix/' + str(FX_UNMUTE_BUS).zfill(2) + ' OFF\n')
        
        # for channels only, set name from additional spreadsheet data
        if NAME_CHANS:
            for chan in range(0, NUM_CHANS):
                name_on_or_above = current_or_previous_channel_name(ods, row_index, chan + FIRST_CHAN_NAME_COL)
                chan_num = string_to_int(read_cell_as_string(ods, PATH_NUM_ROW, chan + FIRST_CHAN_COL))
                if name_on_or_above != '':
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/config/name "' + name_on_or_above + '"\n')
        
        # mute specified channels in range
        if OTHER_MUTES:
            for chan in range(0, OTHER_MUTES_NUM_CHANS):
                mute_this_chan = False
                for col in OTHER_MUTES_COLS:
                    mute_info = read_cell_as_string(ods, row_index, col)
                    if mute_info != '':
                        mute_chan = string_to_int(mute_info)
                        if mute_chan == chan + OTHER_MUTES_FIRST_CHAN:
                            mute_this_chan = True
                if mute_this_chan:
                    snp_file.write('/ch/' + str(chan + OTHER_MUTES_FIRST_CHAN).zfill(2) + '/mix/on OFF\n')
                else:
                    snp_file.write('/ch/' + str(chan + OTHER_MUTES_FIRST_CHAN).zfill(2) + '/mix/on ON\n')
        
        # finally we write out the new DCA labels
        for dca in range(0, NUM_DCAS):
            label = read_cell_as_string(ods, row_index, dca + FIRST_DCA_COL)
            next_label = next_dca_label(ods, row_index, dca + FIRST_DCA_COL)
            if label != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + label + '"\n')
                if DCA_ALT_LABEL_COLORS and label in DCA_ALT_LABELS:
                    snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_ALT_LABEL_COLOR + '\n')
                elif DCA_SAME_ON_NEXT_CUE and label == next_label:
                    snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_SAME_ON_NEXT_CUE_COLOR + '\n')
                else:
                    snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_COLOR + '\n')
            elif DCA_ACTIVE_ON_NEXT_CUE and next_label != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + next_label + '"\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_ACTIVE_ON_NEXT_CUE_COLOR + '\n')
            else:
                snp_file.write('/dca/' + str(dca + 1) + '/config/name ""\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color OFF\n')
        
        # close snippet file
        snp_file.close()
        
        # store
        cue_numbers.append(cue_number)
        cue_labels.append(cue_label)
        
        # next
        snp_index = snp_index + 1
    
    #
    # generate show file
    #
    
    # report
    print('Creating show file...')
    
    # open show file
    shw_file = open(show_name + '.shw', 'w')
    
    # start show
    shw_file.write('#2.6#\n')
    shw_file.write('show "' + show_name +'" 0 0 0 0 0 0 0 0 0 0 "X32-Edit 3.00"\n')
    
    # report
    print('Writing cues...')
    
    # write cues
    cue_index = 0
    for cue_number in cue_numbers:
        shw_file.write('cue/' + str(cue_index).zfill(3) + ' ' + cue_number + ' "' + cue_labels[cue_index] + '" 0 -1 ' + str(cue_index) + ' 0 1 0 0\n')
        cue_index = cue_index + 1
    
    # report
    print('Writing snippets...')
    
    # write snippet refs
    snp_index = 0
    for cue_number in cue_numbers:
        shw_file.write('snippet/' + str(snp_index).zfill(3) + ' "' + cue_labels[snp_index] + '" 0 0 0 0 1\n')
        snp_index = snp_index + 1
        
    # close show file
    shw_file.close()
    
    # all done
    print('Done!')
