#!/usr/bin/python

# NB: As per Debian Python policy #!/usr/bin/env python2 is not used here.

"""
sel-settings-terminal.py

A tool to (at least) extract information from Transpower setting 
spreadsheets.

Usage defined by running with option -h.

This tool can be run from the IDLE prompt using the main def.

Thoughtful ideas most welcome. 

Installation instructions (for Python *2.7.9*):
 - pip install xlrd
 - pip install tablib

 - or if behind a proxy server: pip install --proxy="user:password@server:port" packagename
 - within Transpower: pip install --proxy="transpower\mulhollandd:password@tptpxy001.transpower.co.nz:8080" tablib    
 
 C:\Python27\Scripts\pip install --proxy="transpower\mulhollandd:nottelling@tptpxy001.transpower.co.nz:8080" xlrd
 
TODO: 
 - so many things
 - sorting options on display and dump output?    
 - sort out guessing of Transpower standard design version 
 - sort out dumping all parameters and argparse dependencies
 - sort out extraction of DNP data
"""

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ["Kenneth Reitz https://github.com/kennethreitz/tablib"]
__license__ = "GPL"
__version__ = '0.02'
__maintainer__ = "Daniel Mulholland"
__hosted__ = "https://github.com/danyill/sel-settings-terminal/"
__email__ = "dan.mulholland@gmail.com"

# update this line if using from Python interactive
#__file__ = r'W:\Education\Current\pytooldev\tp-setting-excel-tool'

import sys
import os
import argparse
import glob
import regex
import tablib
import xlrd
import string

BASE_PATH = os.path.dirname(os.path.realpath(__file__))
OUTPUT_FILE_NAME = "output"
OUTPUT_HEADERS = ['Filename','Setting Name','Val','Spreadsheet Reference']
EXCEL_FILE_REGEX = '(xls|xlsx|xlsm)$'
TXT_EXTENSION = 'TXT'
SEL_EXPRESSION = r'[\w :+/\\()!,.\-_\\*]*'
SEL_SETTING_EOL = r'\x1c\r\n'
SEL_SETTING_NAME = r'[\w _]*'
SEL_FID_EXPRESSION='^FID=([\w :+/\\()!,.\-_\\*]*)\r\n'

SEL_EXPRESSION_RETURN = r'[\w :+/\\()!,.\-_\\*]'
SEL_SETTING_NAME_RETURN = r'([A-Z0-9 _]{6})='
SEL_FID_EXPRESSION='^FID=([\w :+/\\()!,.\-_\\*]*)\r\n'

PARAMETER_SEPARATOR = ':'

SEL_SEARCH_EXPR = {\
    'G1': [['Group 1\r\nGroup Settings:', 'SELogic group 1\r\n'], \
           ['=\>', 'Group [23456]\r\nGroup Settings:', 'SELogic group [23456]\r\n'] \
          ], \
    'G2': [['Group 2\r\nGroup Settings:', 'SELogic group 2\r\n'], \
           ['=\>', 'Group [13456]\r\nGroup Settings:', 'SELogic group [13456]\r\n'] \
          ], \
    'G3': [['Group 3\r\nGroup Settings:', 'SELogic group 3\r\n'], \
           ['=\>', 'Group [12456]\r\nGroup Settings:', 'SELogic group [12456]\r\n'] \
          ], \
    'G4': [['Group 4\r\nGroup Settings:', 'SELogic group 4\r\n'], \
           ['=\>', 'Group [12356]\r\nGroup Settings:', 'SELogic group [12356]\r\n'] \
          ],\
    'G5': [['Group 5\r\nGroup Settings:', 'SELogic group 5\r\n'], \
           ['=\>', 'Group [12346]\r\nGroup Settings:', 'SELogic group [12346]\r\n'] \
          ], \
    'G6': [['Group 6\r\nGroup Settings:', 'SELogic group 6\r\n'], \
           ['=\>', 'Group [12345]\r\nGroup Settings:', 'SELogic group [12345]\r\n'] \
          ], \
    'P1': [['Port 1\r\n'], ['$', '=\>', 'Port [2345F]\r\n']], \
    'P2': [['Port 2\r\n'],['$', '=\>', 'Port [2345F]\r\n']], \
    'P3': [['Port 3\r\n'],['$', '=\>', 'Port [2345F]\r\n']], \
    'PF': [['Port F\r\n'],['$', '=\>', 'Port [2345F]\r\n']], \
    }

OUTPUT_HEADERS = ['File','Setting Name','Val']


def main(arg=None):
    parser = argparse.ArgumentParser(
        description='Process individual or multiple RDB files and produce summary'\
            ' of results as a csv or xls file.',
        epilog='Enjoy. Bug reports and feature requests welcome. Feel free to build a GUI :-)',
        prefix_chars='-/')

    parser.add_argument('-o', choices=['csv','xlsx'],
                        help='Produce output as either comma separated values (csv) or as'\
                        ' a Micro$oft Excel .xls spreadsheet. If no output provided then'\
                        ' output is to the screen.')

    parser.add_argument('path', metavar='PATH|FILE', nargs=1, 
                       help='Go recursively go through path PATH. Redundant if FILE'\
                       ' with extension .rdb is used. When recursively called, only'\
                       ' searches for files with:' +  TXT_EXTENSION + '. Globbing is'\
                       ' allowed with the * and ? characters.')

    parser.add_argument('-s', '--screen', action="store_true",
                       help='Show output to screen')

    parser.add_argument('-a', '--all', action="store_true",
                       help='Output all settings!')                       
                       
    # Not implemented yet
    #parser.add_argument('-d', '--design', action="store_true",
    #                   help='Attempt to determine Transpower standard design version and' \
    #                   ' include this information in output')
                       
    parser.add_argument('settings', metavar='G:S', type=str, nargs='+',
                       help='Settings in the form of G:S where G is the group'\
                       ' and S is the SEL variable name. If G: is omitted the search' \
                       ' goes through all groups. Otherwise G should be the '\
                       ' group of interest. S should be the setting name ' \
                       ' e.g. OUT201.' \
                       ' Examples: G1:50P1P or G2:50P1P or 50P1P' \
                       ' '\
                       ' You can also get port settings using P:S'
                       ' Note: Applying a group for a non-grouped setting is unnecessary'\
                       ' and will prevent you from receiving results.'\
                       ' Special parameters are the following self-explanatory items:'\
                       ' FID, PARTNO, DEVID')

    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)

    if arg == None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arg.split())
    
    # read in list of files
    files_to_do = return_file_paths(args.path, TXT_EXTENSION)
    
    # sort out the reference data for knowing where to search in the text string
    lookup = SEL_SEARCH_EXPR
    if files_to_do != []:
        process_txt_files(files_to_do, args, lookup)
    else:
        print('Found nothing to do for path: ' + args.path[0])
        sys.exit()
        os.system("Pause")
    
def return_file_paths(args_path, file_extension):
    paths_to_work_on = []
    for p in args_path:
        p = p.translate(None, ",\"")
        if not os.path.isabs(p):
            paths_to_work_on +=  glob.glob(os.path.join(BASE_PATH,p))
        else:
            paths_to_work_on += glob.glob(p)
            
    files_to_do = []
    # make a list of files to iterate over
    if paths_to_work_on != None:
        for p_or_f in paths_to_work_on:
            if os.path.isfile(p_or_f) == True:
                # add file to the list
                print os.path.normpath(p_or_f)
                files_to_do.append(os.path.normpath(p_or_f))
            elif os.path.isdir(p_or_f) == True:
                # walk about see what we can find
                files_to_do = walkabout(p_or_f, file_extension)
    return files_to_do        

def walkabout(p_or_f, file_extension):
    """ searches through a path p_or_f, picking up all files with EXTN
    returns these in an array.
    """
    return_files = []
    for root, dirs, files in os.walk(p_or_f, topdown=False):
        #print files
        for name in files:
            if (os.path.basename(name)[-3:]).upper() == file_extension:
                return_files.append(os.path.join(root,name))
    return return_files
    
def process_txt_files(files_to_do, args, reference_data):
    parameter_info = []
        
    for filename in files_to_do:      
        extracted_data = extract_parameters(filename, args.settings, reference_data)
        parameter_info += extracted_data

    # for exporting to Excel or CSV
    data = tablib.Dataset()    
    for k in parameter_info:
        data.append(k)
    data.headers = OUTPUT_HEADERS

    # don't overwrite existing file
    name = OUTPUT_FILE_NAME 
    if args.o == 'csv' or args.o == 'xlsx': 
        # this is stupid and klunky
        while os.path.exists(name + '.csv') or os.path.exists(name + '.xlsx'):
            name += '_'        

    # write data
    if args.o == None:
        pass
    elif args.o == 'csv':
        with open(name + '.csv','wb') as output:
            output.write(data.csv)
    elif args.o == 'xlsx':
        with open(name + '.xlsx','wb') as output:
            output.write(data.xlsx)

    if args.screen == True:
        display_info(parameter_info)

def extract_parameters(filename, settings, reference_data):
    parameter_info=[]

    # read data
    with open(filename,'r') as f:
        read_data = f.read()

    """
    How this regex works:
     * (\r\n| |^)
       - Looks for either a new line CR/LF  or a space or the start of the file.
       - This is always true in process terminal views.
     
     * ([A-Z0-9 _]{6})
       - SEL setting names are typically uppercase without spaces comprising 
         characters A-Z 0-9 and sometimes with underscores (exception, DNP)
     
     * =
       - Then followed by an equals character
     
     * (?>([\w :+/\\()!,.\-_\\*]+)
       - There's quite a few options for what can be in a SEL expression
       - This probably doesn't take them all into account add to suit 
       - This is an atomic group expression which is a solution for making 
         sure the delimiter doesn't get "eaten" because the delimiter is 
         comprised of the same characters as the expression.
         
       - This is well described here: http://www.rexegg.com/regex-quantifiers.html
     
     * ([ ]{0}[A-Z0-9 _]{6}=|\r\n
       - Then the delimiter comes. This is the next SEL setting name, if there
         are multiple columns. Alternatively the delimiter is a newline CR/LF
         combination.
    """

    """
    TODO: This is how the --all or -a parameter should be implemented
    results = regex.findall('(\r\n| |^)([A-Z0-9 _]{6})=(?>([\w :+/\\()!,.\-_\\*]+)([ ]{0}[A-Z0-9 _]{6}=|\r\n))', 
        data, flags=regex.MULTILINE, overlapped=True)
    
    Just need to break down the groups. Trivial. Execise for the reader.
    :-)
    """
    
    for parameter in settings:
        data = read_data
        # parameter is e.g. G1:50P1P and there is a separator
        # if parameter.find(PARAMETER_SEPARATOR) != -1:
        if parameter.find(PARAMETER_SEPARATOR) != -1:
            grouper = parameter.split(PARAMETER_SEPARATOR)[0]
            setting = parameter.split(PARAMETER_SEPARATOR)[1]
            
        if parameter.find(PARAMETER_SEPARATOR) == -1 or \
            SEL_SEARCH_EXPR[grouper] == None:
            # print 'Searching the whole file without bounds'
            if parameter in ['FID', 'PARTNO', 'DEVID']:
                result = get_special_parameter(parameter,data)        
            else:
                result = find_SEL_text_parameter(setting, [data])
        
        else:
            # now search inside this data group for the desired setting
            data = find_between_text( \
                start_options = SEL_SEARCH_EXPR[grouper][0], \
                end_options = SEL_SEARCH_EXPR[grouper][1],  
                text = data) 
        
            if data:
                result = find_SEL_text_parameter(setting, data)

        if result <> None:
            filename = os.path.basename(filename)
            parameter_info.append([filename, parameter, result])
            
    return parameter_info

def find_SEL_text_parameter(setting, data_array):
    
    for r in data_array:
        # Example for TR setting: 
        #  - (\r\n| |^)(TR    )=(?>([\w :+/\\()!,.\-_\\*]+)([ ]{0}[A-Z0-9 _]{6}=|\r\n))
        found_parameter = regex.findall('(\r\n| |^)(' + \
                string.ljust(setting, 6, ' ') + \
                ')=(?>([\w :+/\\()!,.\-_\\*]+)([ ]{0}[A-Z0-9 _]{6}=|\r\n))', \
                r, flags=regex.MULTILINE, overlapped=True)
         
        if found_parameter:
            return found_parameter[0][2]
        
def find_between_text(start_options, end_options, text):
    # return matches between arbitrary start and end options
    # with matches across lines
    results = []
    start_regex = ''
    for k in start_options:
        start_regex = k 
        
        # create ending regex expression
        end_regex = '('                    
        for k in end_options:
            end_regex += k + '|'
        end_regex = end_regex[0:-1]                    
        end_regex += ')'
        
        result = regex.findall(start_regex + '((.|\n)+?)' + end_regex, text, flags = regex.MULTILINE)
        
        if result:
            # print result[0][0]
            results.append(result[0][0])
        
    return results

def get_special_parameter(name,data):
    # Something like:
    # name=FID for "FID=SEL-351S-6-R107-V0-Z003003-D20011129","0958"
    # name=PARTNO for "PARTNO=0351S61H3351321","05AE"
    # name=DEVID for "DEVID=TMU 2782","0402"
    return regex.findall(r'^\"' + name + r'=([\w :+/\\()!,.\-_\\*\"]*\r\n)', 
        data, flags=regex.MULTILINE, overlapped=True)

def get_dnp(name, data):
    # Not implemented yet
    # Analogs  = 0 2 4 8 10 12 31 35 106 
    # Binaries = 295 677 678 223 216 224 1020 1021 1022 296 527 571 567 595 735  \
    #       734 233 242 740 251 179 180 181 360 361 362 863 364 865 866 867  \
    #       767 766 765 679 680 681 864 
    pass

def display_info(parameter_info):
    lengths = []
    # first pass to determine column widths:
    for line in parameter_info:
        for index,element in enumerate(line):
            try:
                lengths[index] = max(lengths[index], len(element))
            except IndexError:
                lengths.append(len(element))
    
    parameter_info.insert(0,OUTPUT_HEADERS)
    # now display in columns            
    for line in parameter_info:
        display_line = '' 
        for index,element in enumerate(line):
            display_line += element.ljust(lengths[index]+2,' ')
        print display_line
        
if __name__ == '__main__': 
    if len(sys.argv) == 1 :
        main(r'--all -o csv in G1:TID FID G1:TR G1:81D1P G1:81D1D G1:81D2P G1:81D2P G1:E81')           
    else:
        main()
    os.system("Pause")
        
"""
def test_xls():

    sys.stdout = open('log', 'w')
    
    my_file = return_file_paths(['in'], 'XLS')[0]
    
    workbook = xlrd.open_workbook(my_file)
    print workbook.sheet_names()
    # worksheet = workbook.sheet_by_name('Settings_Printout')
    
    worksheets = workbook.sheet_names()
    for worksheet_name in worksheets:
        worksheet = workbook.sheet_by_name(worksheet_name)
        print worksheet.name
    
    workksheet = worksheets[0]
    #worksheet = workbook.sheet_by_name('Group_1')
    worksheet = workbook.sheet_by_name('Settings_Printout')
    num_rows = worksheet.nrows - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        print row
        
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        print 'Row:', curr_row
        curr_cell = -1
        while curr_cell < num_cells:
            curr_cell += 1
            # Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            print '	', cell_type, ':', cell_value
"""
