#!/usr/bin/env python

################################################################################
#
# X32 Snippets
#
# Last mod:
# January 23, 2017
#
# Written by:
# Simon Eves (simon@eves.us)
#
# Additional contributions by:
# Craig Flint (craig@cliveflint.co.uk)
# Aaron Yoffe (aaron@blownfuse.org)
#
# Free for non-commercial use
#
################################################################################

VERSION = "1.3"

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

SHEET_NAME      = "Sheet1" # spreadsheet sub-sheet name containing all data

SKIP_ROWS       = 2      # number of spreadsheet rows to skip before extracting data

SNIPPET_COL     = 0      # spreadsheet column of snippet index data (must be monotonically incrementing integers)
CUE_COL         = 1      # spreadsheet column of cue data (any short string)

FIRST_CHAN      = 1      # number of first physical channel to control
FIRST_CHAN_COL  = 2      # spreadsheet column of data for that first channel
NUM_CHANS       = 23     # number of contiguous set of channels

FIRST_BUS       = 0      # number of first physical bus to control (13 = FX1)
FIRST_BUS_COL   = 0      # spreadsheet column of data for that first bus
NUM_BUSES       = 0      # number of contiguous set of buses (4 = FX1-4)

FIRST_AUXIN     = 1      # number of first physical auxin to control (5 = AuxIn 5)
FIRST_AUXIN_COL = 25     # spreadsheet column of data for that first auxin
NUM_AUXINS      = 6      # number of contiguous set of auxins (2 = AuxIn 5-6)

FIRST_DCA_COL   = 32     # spreadsheet column of data for first DCA
NUM_DCAS        = 8      # number of DCAs

WARN_DCAS       = False  # show next cue's active DCAs in red?

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
        return cell
    except:
        return ''

def process_paths(ods, snp_file, row_index, first_path, first_path_col, num_paths, osc_prefix):

    # general function to process paths of any of the three types

    # first we find which paths in this cue have any DCA assignment and therefore need to be unmuted
    mutes = []
    for path in range(0, num_paths):
        dca = str(ods_cell(ods, row_index, path + first_path_col))
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
        dca = str(ods_cell(ods, row_index, path + first_path_col))
        bitmap = 0
        if dca != '':
            bitmap = 1 << (int(float(dca)) - 1)
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
    
    # read the file
    ods = get_data(ods_file_name)[SHEET_NAME]
    
    # init these
    row_index = SKIP_ROWS
    snp_index = 0
    
    # iterate rows until the end
    for row_index in range(SKIP_ROWS, len(ods)):

        # get snippet
        snippet = ods_cell(ods, row_index, SNIPPET_COL)
        if type(snippet) == float:
            snippet = str(int(snippet))
        else:
            snippet = str(snippet)
        
        # skip rows with no snippet
        if snippet == '':
            continue
            
        # get cue
        cue = str(ods_cell(ods, row_index, CUE_COL))
        
        # report snippet
        print "Generating new Snippet '" + snippet + "', Cue '" + cue + "'"
            
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

        # finally we write out the new DCA labels
        for dca in range(0, NUM_DCAS):
            label = str(ods_cell(ods, row_index, dca + FIRST_DCA_COL))
            labelbelow = str(ods_cell(ods, row_index + 1, dca + FIRST_DCA_COL))
            if label != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + label + '"\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color GN\n')
            elif WARN_DCAS && labelbelow != '':
                snp_file.write('/dca/' + str(dca + 1) + '/config/name "' + labelbelow + '"\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color RD\n')
            else:
                snp_file.write('/dca/' + str(dca + 1) + '/config/name ""\n')
                snp_file.write('/dca/' + str(dca + 1) + '/config/color OFF\n')

        # close file
        snp_file.close()

        # the end?
        if snippet == "END":
            break
            
        # next
        snp_index = snp_index + 1
    
    # all done
    print "Goodbye!"
