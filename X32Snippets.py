#!/usr/bin/env python

################################################################################
#
# X32 Snippets
#
# Last mod:
# July 16, 2017
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

VERSION = "1.6"

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

# values for Heathers

SHEET_NAME             = 'Sheet1' # spreadsheet sub-sheet name containing all data

SKIP_ROWS              = 6      # number of spreadsheet rows to skip before extracting data

CUE_NUM_COL            = 0      # spreadsheet column of cue number data (can be integer or tenths, e.g. "1" or "1.1")
CUE_LABEL_COL          = 1      # spreadsheet column of cue data (any short string)

PATH_NUM_ROW           = 0      # row containing path (chan, bus, aux) numbers

FIRST_CHAN_COL         = 4      # spreadsheet column of data for that first channel
NUM_CHANS              = 17     # number of contiguous set of channels

FIRST_BUS_COL          = 0      # spreadsheet column of data for that first bus
NUM_BUSES              = 0      # number of contiguous set of buses (4 = FX1-4)

FIRST_AUXIN_COL        = 0      # spreadsheet column of data for that first auxin
NUM_AUXINS             = 0      # number of contiguous set of auxins (2 = AuxIn 5-6)

FIRST_DCA_COL          = 22     # spreadsheet column of data for first DCA
NUM_DCAS               = 7      # number of DCAs
DCA_COLOR              = 'WH'   # color for active DCA labels

# Craig Flint warn upcoming DCAs
CF_WARN_DCAS           = False  # show next cue's active DCAs in red?
CF_WARN_COLOR          = 'RD'   # color for warning DCA labels (if WARN_DCAS)

# Craig Flint name channels
CF_NAME_CHANS          = False # set channel names
CF_FIRST_CHAN_NAME_COL = 0     # spreadsheet column of name for the first channel

# Craig Flint alternative color DCA labels
CF_ALT_LABEL_COLORS    = False
CF_ALT_LABELS          = [ 'Reverb', 'Sh Dly', 'Lg Dly', 'Delay' ] # color these DCA labels differently
CF_ALT_COLOR           = 'MG'

# Simon Eves FX send automation
SE_FX                  = True   # negative path DCA indexes mean also un-mute that path's FX send to the bus below
SE_FX_BUS              = '14'   # set the send from the paths to this FX bus (zero-padded number as string)

# Simon Eves band muting automation
SE_BAND_MUTES          = False  # mute one band channel per cue, all others will be unmuted (e.g. for switching basses)
SE_FIRST_BAND_CHAN     = 17     # first board channel for the band
SE_NUM_BAND_CHANS      = 10     # number of consecutive board channels for the band
SE_BAND_MUTES_COL      = 41     # spreadsheet column of data for which channel to mute


################################################################################
# Functions
################################################################################

def ods_cell(d, r, c):

    # robust function to pull contents of a spreadsheet cell as a string, regardless of actual contents

    try:
        row = d[r]
        cell = row[c]
        #print "DEBUG: cell " + str(r) + "/" + str(c) + " is type " + str(type(cell))
        if type(cell) == unicode:
            cell = cell.encode('utf-8')
        return str(cell)
    except:
        return ''

def process_paths(ods, snp_file, row_index, first_path_col, num_paths, osc_prefix):

    # general function to process paths of any of the three types
    # we pull the actual board path number from PATH_NUM_ROW

    # first we find which paths in this cue have any DCA assignment and therefore need to be unmuted
    mutes = []
    for path in range(0, num_paths):
        dca = ods_cell(ods, row_index, path + first_path_col)
        if dca != '':
            mutes.append(False)
        else:
            mutes.append(True)
            
    # then we write out mute-ons for paths which have become muted
    for path in range(0, num_paths):
        if mutes[path]:
            path_num = int(float(ods_cell(ods, PATH_NUM_ROW, path + first_path_col)))
            snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/mix/on OFF\n')

    # then we write out the new DCA assignments
    for path in range(0, num_paths):
        dca = ods_cell(ods, row_index, path + first_path_col)
        bitmap = 0
        if dca != '':
            bitmap = 1 << (abs(int(float(dca))) - 1)
        else:
            bitmap = 0
        path_num = int(float(ods_cell(ods, PATH_NUM_ROW, path + first_path_col)))
        snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/grp/dca ' + str(bitmap) + '\n')
    
    # then we write out mute-offs for paths which have become un-muted
    for path in range(0, num_paths):
        if not mutes[path]:
            path_num = int(float(ods_cell(ods, PATH_NUM_ROW, path + first_path_col)))
            snp_file.write('/' + osc_prefix + '/' + str(path_num).zfill(2) + '/mix/on ON\n')
            
def next_dca_label(ods, row_index, col):

    # function to find the contents of the cell in col 'col' in the next row below 'row_index' that has a valid cue number

    # are we already at the END?
    search_cue = ods_cell(ods, row_index, CUE_NUM_COL)
    if search_cue == 'END':
        return ''

    # find the next row that has a valid cue
    search_row = row_index + 1
    search_cue = ods_cell(ods, search_row, CUE_NUM_COL)
    while search_cue == '':
        search_row = search_row + 1
        search_cue = ods_cell(ods, search_row, CUE_NUM_COL)

    # was it the END
    if search_cue == 'END':
        return ''

    # otherwise return whatever's in the column
    return ods_cell(ods, search_row, col)

def current_or_previous_channel_name(ods, row_index, col):

    # function to find the contents of the next non-empty cell in col 'col' and row 'row_index' or above

    # are we already at the top?
    if row_index <= SKIP_ROWS:
        return ''

    # search upwards to find the previous non-empty cell in this column
    search_row = row_index
    search_name = ods_cell(ods, search_row, col)
    while search_name == '' and search_row > SKIP_ROWS:
        search_row = search_row - 1
        search_name = ods_cell(ods, search_row, col)

    # return
    return search_name


################################################################################
# Main
################################################################################    
        
if __name__ == "__main__":

    title = '# X32 Snippets v' + VERSION
    print '#' * len(title)
    print title
    print '#' * len(title)
    
    #
    # validate command-line parameters
    #
    
    if len(sys.argv) != 3:
        print "";
        print "Usage: X32Snippets.py <ods_file_name> <show_name>"
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
    print "Opening spreadsheet..."
    
    # read the file
    ods = get_data(ods_file_name)[SHEET_NAME]
    
    # init these
    row_index = SKIP_ROWS
    snp_index = 0
    cue_numbers = []
    cue_labels = []
    
    # report
    print "Creating cues..."
    
    # iterate rows until the end
    for row_index in range(SKIP_ROWS, len(ods)):
        
        # get cue
        cue = ods_cell(ods, row_index, CUE_NUM_COL)
        
        # skip rows with no cue
        if cue == '':
            continue
            
        # the end?
        if cue == 'END':
            break
            
        # get cue number
        cue_number = ''
        try:
            cue_number = str(int(float(cue) * 100.0))
        except:
            print 'ERROR: Found invalid cue number at row ' + str(row_index + 1)
            sys.exit()
            
        # get cue label
        cue_label = ods_cell(ods, row_index, CUE_LABEL_COL)
        
        # report snippet
        print 'Generating new cue "' + cue + '", label "' + cue_label + '"'
            
        # open snippet file
        snp_file = open(show_name + '.' + str(snp_index).zfill(3) + '.snp', 'w')
        
        # start snippet
        snp_file.write('#2.1# "' + cue + '" 0 0 0 0 0\n')
        
        # process channels
        process_paths(ods, snp_file, row_index, FIRST_CHAN_COL, NUM_CHANS, 'ch')
        
        # process buses
        process_paths(ods, snp_file, row_index, FIRST_BUS_COL, NUM_BUSES, 'bus')
        
        # process auxins
        process_paths(ods, snp_file, row_index, FIRST_AUXIN_COL, NUM_AUXINS, 'auxin')
        
        # Simon Eves
        # for channels only, also control the mute of the given FX bus send for this path
        if SE_FX:
            for chan in range(0, NUM_CHANS):
                dca = ods_cell(ods, row_index, chan + FIRST_CHAN_COL)
                fx_on = False
                if dca != '':
                    fx_on = int(float(dca)) < 0
                chan_num = int(float(ods_cell(ods, PATH_NUM_ROW, chan + FIRST_CHAN_COL)))
                if fx_on:
                    print 'DEBUG: activating FX send on channel ' + str(chan_num)
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/mix/' + SE_FX_BUS + ' ON\n')
                else:
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/mix/' + SE_FX_BUS + ' OFF\n')
        
        # Craig Flint
        # for channels only, set name from additional spreadsheet data
        if CF_NAME_CHANS:
            for chan in range(0, NUM_CHANS):
                name_on_or_above = current_or_previous_channel_name(ods, row_index, chan + CF_FIRST_CHAN_NAME_COL)
                chan_num = int(float(ods_cell(ods, PATH_NUM_ROW, chan + FIRST_CHAN_COL)))
                if name_on_or_above != '':
                    print 'DEBUG: renaming channel ' + str(chan_num) + ' as "' + name_on_or_above + '"'
                    snp_file.write('/ch/' + str(chan_num).zfill(2) + '/config/name "' + name_on_or_above + '"\n')
        
        # Simon Eves
        # mute specified band channels
        if SE_BAND_MUTES:
            mute = ods_cell(ods, row_index, SE_BAND_MUTES_COL)
            if mute != '':
                mute_chan = int(float(mute))
                for chan in range(0, SE_NUM_BAND_CHANS):
                    if (chan + SE_FIRST_BAND_CHAN) == mute_chan:
                        print 'DEBUG: muting band channel ' + str(chan + SE_FIRST_BAND_CHAN)
                        snp_file.write('/ch/' + str(chan + SE_FIRST_BAND_CHAN).zfill(2) + '/mix/on OFF\n')
                    else:
                        snp_file.write('/ch/' + str(chan + SE_FIRST_BAND_CHAN).zfill(2) + '/mix/on ON\n')
        
        # finally we write out the new DCA labels
        for dca in range(0, NUM_DCAS):
            label = ods_cell(ods, row_index, dca + FIRST_DCA_COL)
            label_below = next_dca_label(ods, row_index, dca + FIRST_DCA_COL)
            if label != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + label + '"\n')
                if CF_ALT_LABEL_COLORS and label in CF_ALT_LABELS:
                    snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + CF_ALT_COLOR + '\n')
                else:
                    snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_COLOR + '\n')
            elif CF_WARN_DCAS and label_below != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + label_below + '"\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + CF_WARN_COLOR + '\n')
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
    print 'Creating show file...'
    
    # open show file
    shw_file = open(show_name + '.shw', 'w')
    
    # start show
    shw_file.write('#2.6#\n')
    shw_file.write('show "' + show_name +'" 0 0 0 0 0 0 0 0 0 0 "X32-Edit 3.00"\n')
    
    # report
    print 'Writing cues...'
    
    # write cues
    cue_index = 0
    for cue_number in cue_numbers:
        shw_file.write('cue/' + str(cue_index).zfill(3) + ' ' + cue_number + ' "' + cue_labels[cue_index] + '" 0 -1 ' + str(cue_index) + ' 0 1 0 0\n')
        cue_index = cue_index + 1
    
    # report
    print 'Writing snippets...'
    
    # write snippet refs
    snp_index = 0
    for cue_number in cue_numbers:
        shw_file.write('snippet/' + str(snp_index).zfill(3) + ' "' + cue_labels[snp_index] + '" 0 0 0 0 1\n')
        snp_index = snp_index + 1
        
    # close show file
    shw_file.close()
    
    # all done
    print 'Done!'
