#!/usr/bin/env python

################################################################################
#
# X32 Snippets
#
# Last mod:
# May 10, 2017
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

VERSION = "1.5 Beta 1"

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

SHEET_NAME             = "Sheet1" # spreadsheet sub-sheet name containing all data

SKIP_ROWS              = 5      # number of spreadsheet rows to skip before extracting data

CUE_NUM_COL            = 0      # spreadsheet column of snippet index data (must be monotonically incrementing integers)
CUE_LABEL_COL          = 1      # spreadsheet column of cue data (any short string)

FIRST_CHAN             = 1      # number of first physical channel to control
FIRST_CHAN_COL         = 4      # spreadsheet column of data for that first channel
NUM_CHANS              = 21     # number of contiguous set of channels

FIRST_BUS              = 0      # number of first physical bus to control (13 = FX1)
FIRST_BUS_COL          = 0      # spreadsheet column of data for that first bus
NUM_BUSES              = 0      # number of contiguous set of buses (4 = FX1-4)

FIRST_AUXIN            = 0      # number of first physical auxin to control (5 = AuxIn 5)
FIRST_AUXIN_COL        = 0      # spreadsheet column of data for that first auxin
NUM_AUXINS             = 0      # number of contiguous set of auxins (2 = AuxIn 5-6)

FIRST_DCA_COL          = 26     # spreadsheet column of data for first DCA
NUM_DCAS               = 8      # number of DCAs
DCA_COLOR              = 'WH'   # color for active DCA labels

# Craig Flint warn upcoming DCAs
CF_WARN_DCAS           = False  # show next cue's active DCAs in red?
CF_WARN_COLOR          = 'RD'   # color for warning DCA labels (if WARN_DCAS)

# Craig Flint name channels
CF_NAME_CHANS          = False  # set channel names
CF_FIRST_CHAN_NAME_COL = 35     # spreadsheet column of name for the first channel

# STC FX send automation
STC_FX                 = False  # negative path DCA indexes mean also un-mute that path's FX send to the bus below
STC_FX_BUS             = '14'   # set the send from the paths to this FX bus (zero-padded number as string)


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

def process_paths(ods, snp_file, row_index, first_path, first_path_col, num_paths, osc_prefix):

    # general function to process paths of any of the three types

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
            snp_file.write('/' + osc_prefix + '/' + str(path + first_path).zfill(2) + '/mix/on OFF\n')

    # then we write out the new DCA assignments
    for path in range(0, num_paths):
        dca = ods_cell(ods, row_index, path + first_path_col)
        bitmap = 0
        if dca != '':
            bitmap = 1 << (abs(int(float(dca))) - 1)
        else:
            bitmap = 0
        snp_file.write('/' + osc_prefix + '/' + str(path + first_path).zfill(2) + '/grp/dca ' + str(bitmap) + '\n')
    
    # then we write out mute-offs for paths which have become un-muted
    for path in range(0, num_paths):
        if not mutes[path]:
            snp_file.write('/' + osc_prefix + '/' + str(path + first_path).zfill(2) + '/mix/on ON\n')
            

################################################################################
# Main
################################################################################    
        
if __name__ == "__main__":

    print "#####################"
    print "# X32 Snippets v" + VERSION
    print "#####################"
    
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
        if cue == "END":
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
        process_paths(ods, snp_file, row_index, FIRST_CHAN, FIRST_CHAN_COL, NUM_CHANS, 'ch')
        
        # process buses
        process_paths(ods, snp_file, row_index, FIRST_BUS, FIRST_BUS_COL, NUM_BUSES, 'bus')
        
        # process auxins
        process_paths(ods, snp_file, row_index, FIRST_AUXIN, FIRST_AUXIN_COL, NUM_AUXINS, 'auxin')

        # STC
        # for channels only, also control the mute of the given FX bus send for this path
        if STC_FX:
            for chan in range(0, NUM_CHANS):
                dca = ods_cell(ods, row_index, chan + FIRST_CHAN_COL)
                fx_on = False
                if dca != '':
                    fx_on = int(float(dca)) < 0
                if fx_on:
                    print 'DEBUG: activating FX send on channel ' + str(chan + FIRST_CHAN)
                    snp_file.write('/ch/' + str(chan + FIRST_CHAN).zfill(2) + '/mix/' + STC_FX_BUS + ' ON\n')
                else:
                    snp_file.write('/ch/' + str(chan + FIRST_CHAN).zfill(2) + '/mix/' + STC_FX_BUS + ' OFF\n')
        
        # Craig Flint
        # for channels only, set name from additional spreadsheet data
        # this is incomplete, as it breaks the rules about random-access, but it'll do as a first test
        if CF_NAME_CHANS:
            for chan in range(0, NUM_CHANS):
                name = ods_cell(ods, row_index, chan + CF_FIRST_CHAN_NAME_COL)
                if name != '':
                    print 'DEBUG: renaming channel ' + str(chan + FIRST_CHAN) + ' as "' + name + '"'
                    snp_file.write('/ch/' + str(chan + FIRST_CHAN) + '/config/name "' + name + '"\n')
        
        # finally we write out the new DCA labels
        for dca in range(0, NUM_DCAS):
            label = ods_cell(ods, row_index, dca + FIRST_DCA_COL)
            labelbelow = ods_cell(ods, row_index + 1, dca + FIRST_DCA_COL)
            if label != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + label + '"\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color ' + DCA_COLOR + '\n')
            elif CF_WARN_DCAS and labelbelow != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + labelbelow + '"\n')
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
    print "Creating show file..."
    
    # open show file
    shw_file = open(show_name + '.shw', 'w')
    
    # start show
    shw_file.write('#2.6#\n')
    shw_file.write('show "' + show_name +'" 0 0 0 0 0 0 0 0 0 0 "X32-Edit 3.00"\n')
    
    # report
    print "Writing cues..."
    
    # write cues
    cue_index = 0
    for cue_number in cue_numbers:
        shw_file.write('cue/' + str(cue_index).zfill(3) + ' ' + cue_number + ' "' + cue_labels[cue_index] + '" 0 -1 ' + str(cue_index) + ' 0 1 0 0\n')
        cue_index = cue_index + 1
    
    # report
    print "Writing snippets..."
    
    # write snippet refs
    snp_index = 0
    for cue_number in cue_numbers:
        shw_file.write('snippet/' + str(snp_index).zfill(3) + ' "' + cue_labels[snp_index] + '" 0 0 0 0 1\n')
        snp_index = snp_index + 1
        
    # close show file
    shw_file.close()
    
    # all done
    print "Goodbye!"
