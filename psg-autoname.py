# Autoname GUI and PSG testbed

import string, os, sys, subprocess, threading
from random import shuffle
from glob import glob
from configparser import ConfigParser
import PySimpleGUI as sg

_LOCS = {'SCAN_DIR' : '', 'OUTPUT_DIR' : '', 'WINRAR_PATH': ''} # filled in by config parser and .ini file
NUMBOXES = 5 # number of editing text boxes/segments used by the UI

class Book:
    def __init__(self, filepath):
        if filepath:
            self.filepath = filepath                    # the full file path - usually doesn't change
                                                        # because it's used to move/rename the original file
            self.dirname = os.path.dirname(filepath)    # directory name
            splitname = os.path.splitext(filepath)
            self.name = os.path.basename(splitname[0])  # just the book name, no ext
            self.ext = splitname[1].lower()             # the file extension, convert to lower for safety
            self.filename = self.name + self.ext        # name + file extension
            self.size = self.get_size_str()
            self.seglist = [str.strip(x) for x in self.name.split(' - ')] # segment list
        else:
            self.filepath = self.name = self.ext = self.filename = self.size = ''
            self.seglist = []

    def __repr__(self):
        return self.filepath

    def reassemble_segs(self):
        # if a segment has been added or deleted, this funct fixes the book name to match
        self.name = ' - '.join([x for x in self.seglist])
        self.filename = self.name + self.ext

    def edit_seg(self, segnum, text):
        self.seglist[segnum] = text
        self.reassemble_segs()

    def split_seg(self, window, segnum):
        if segnum < len(self.seglist):  # segnum is 0 based
            splwin = sg.Window('Splitting segment...', 
                               [[sg.Text('Split at text:'), sg.InputText(focus=True)], 
                               [sg.Push(), sg.OK(size=(10, 1)), sg.Cancel(size=(10, 1))]],
                               grab_anywhere=True)
            values = splwin.read()[1]
            splwin.close()
            split = values[0]

            if split:
                before, spl, after = self.seglist[segnum].partition(split)
                self.seglist.insert(segnum, before.strip())
                self.seglist[segnum+1] = spl + after
                self.reassemble_segs()
        else:
            update_statustxt(window, 'Invalid segment number.')

    def get_size_int(self):
        try:
            fsize = round(os.path.getsize(self.filepath) / 1024)
        except Exception as err:
            print(f'Error getting file size {err}')
            return None
        else:
            return fsize

    def get_size_str(self):
        fsize = self.get_size_int()
        if fsize:
            return f'{fsize} KB' if fsize < 1024 else f'{round(fsize/1024, 2)} MB'

    def swap_segs(self, window, x, y):
        # swap positions of segments x and y
        if x <= len(self.seglist) and y <= len(self.seglist):
            self.seglist[x-1], self.seglist[y-1] = self.seglist[y-1], self.seglist[x-1]
            update_statustxt(window, 'Segments swapped.')
            self.reassemble_segs()
        else:
            update_statustxt(window, 'Segment number out of range.')

    def add_seg(self, text=None, pos=None):
        acceptletts = string.ascii_letters + string.digits + "![] ,'.;"

        if not text:
            #newseg = sg.PopupGetText('Text to add:', 'Adding segment...')
            newseg = sg.Window('Adding segment...', 
                               [[sg.Text('Text to add:'), sg.InputText(focus=True)], 
                               [sg.Push(), sg.OK(size=(10, 1)), sg.Cancel(size=(10, 1))]],
                               grab_anywhere=True)
            values = newseg.read()[1]
            newseg.close()
            newseg = values[0]

            # sanitize input
            if newseg:
                newseg = ''.join([x for x in newseg if x and x in acceptletts])
        else:
            newseg = text
        if newseg:
            if not pos: # no position given so append
                self.seglist.append(newseg.strip())
            else: # insert rather than add on the end
                self.seglist.insert(pos, newseg)
            self.capitalize()
            self.reassemble_segs()

    def del_seg(self, window, segnum):
        if len(self.seglist) == 1:
            update_statustxt(window, 'Only one segment remaining, cannot delete.')
            return

        if segnum <= len(self.seglist):
            del self.seglist[segnum-1]
            # now need to reassemble the other bits
            self.reassemble_segs()
        else:
            update_statustxt(window, "Attempt to delete a segment that doesn't exist.")

    def by_replace(self):
        # replace 'by' in book names by splitting
        for x, seg in enumerate(self.seglist):
            seg = seg.replace(' By ', ' by ')
            if ' by ' in seg:
                beforeby, _, afterby = seg.partition(' by ')
                self.seglist[x] = beforeby.strip(' ,-.')
                self.seglist.insert(0, afterby.strip(' ,-.')) # this is probably author so move to front
        self.reassemble_segs()

    def check_title(self, window):
        # runs a couple of checks to catch basic naming errors
        # sanity check for no comma in author's name
        if ',' not in self.seglist[0] and 'Various' not in self.seglist[0]:
            query = 'The author does not seem to have their name reversed.\n\nProceed anyway?\n'
            go = sg.PopupYesNo(query, title='Move file?')
            if go == 'No':
                return False
        if self.name.count('[') != self.name.count(']') or \
           self.name.count('(') != self.name.count(')'):
            update_statustxt(window, "Rename stopped, brackets don't match.")
            return False
        return True

    def finish(self, window, movebook=False):
        newname = _LOCS['OUTPUT_DIR'] + self.filename if movebook else _LOCS['SCAN_DIR'] + self.filename

        while '  ' in newname: #catch sneaky double spaces
            newname = newname.replace('  ', ' ')

        if movebook and self.ext.lower() == '.pdf':
            update_statustxt(window, 'Book is still in PDF format, move halted.')
        elif '40k' in newname or '40K' in newname:
            update_statustxt(window, '"40k" still in book name.')
        elif movebook and self.get_size_int() > 5000:
            update_statustxt(window, 'Book size is over 5MB, move halted.')
        else:
            if self.check_title(window):
                update_statustxt(window, f'Renaming book to {newname}.')
                try:
                    os.rename(self.filepath, newname)
                except Exception as err:
                    update_statustxt(window, f'Error renaming book: {err}')
                else:
                    return True

    def delete(self, window):
        #delete book
        text = f'About to delete {self.filepath}\n\nAre you sure?'
        check =  sg.PopupYesNo(text, title='Delete File?')
        if check == 'Yes':
            update_statustxt(window, "Deleting file...")
            try:
                os.remove(self.filepath)
            except Exception as err:
                update_statustxt(window, f'Error deleting file: {err}')
            else:
                update_statustxt(window, 'File deleted successfully.')
                update_done_txt(window, True) # count a deleted book as done
                return True

    def dupefinder(self, window):
        ignoredwords =  ['The', 'And', 'To', 'Of', 'By', 'With', 'We', 'As']
        #take author's name as first filter, search in subset
        transtable = str.maketrans('', '', ',.&()-[]0123456789')
        # QQQ should check for short auth name like de la Mare
        auth = self.seglist[0].split()[0].translate(transtable) 
        rarlist = [x for x in os.listdir(_LOCS['OUTPUT_DIR']) if x[-4:] == '.rar']
        #rarlist = glob.glob(_LOCS['OUTPUT_DIR'] + '*.rar')
        # note that glob is not case sensitive
        booklist = [os.path.basename(x)[:-4] for x in rarlist if auth.lower() in x.lower()]
        #booklist = [os.path.basename(x)[:-4] for x in glob.glob(f'{_LOCS['OUTPUT_DIR']}\\*{auth.lower()}*.rar')]
        # if book name is only one segment, skip this and just search with author's name
        if len(self.seglist) > 1:
            # otherwise get first word of title, adjusting for series name
            if len(self.seglist) <= 2 or '[' not in self.seglist[1]\
                or "A Very Short Introduction" in self.seglist:
                # for book titles with 1 or 2 segments, like AAA, BBB - CCC
                titlenum = 1
            else:
                # for book titles with 3 or more segments, like Aaa, B - [GGG] - HHH - JJJ, get
                # segment after the segment with ']' in it
                for num, x in enumerate(self.seglist):
                    if ']' in x:
                        titlenum = num + 1
                        break
                else: #fall back to using last segment
                    titlenum = len(self.seglist)- 1

            title = [word for word in self.seglist[titlenum].split()\
                     if len(word) > 2 and word.capitalize() not in ignoredwords]
            if title:
                title = title[0]
            else: # a title like '50 in 50' will cause title list to be empty, so make do with first bit anyway
                title = self.seglist[titlenum].split()[0]
            title = title.translate(transtable)
            if title:
                srchstr = f'Searching on keywords "{auth}" and "{title}": '
                result = [x for x in booklist if title.lower() in x.lower()]
            else:
                srchstr = f'Searching for first word only ("{auth}") as title is too short or invalid.'
                result = booklist
        else:
            srchstr = f'Searching for first word only ("{auth}") as no other keywords found.'
            result = booklist

        if result == []:
            update_statustxt(window, srchstr + 'no matches found.')
            #print("No matches found.")
        else:
            #print(" *** Match Found ***")
            #for x in result: print('>>> ' + x)
            update_statustxt(window, "Matches found.")
            outtext = srchstr + '\n\nMatches found:\n• ' + '\n• '.join([x for x in result])
            if len(result) <= 10:
                sg.PopupOK(outtext)
            else: # stop lots of results overflowing the normal popup window
                sg.PopupScrolled(outtext, size=(70, 12))

    def rar(self, window):
        if self.ext == '.rar':
            update_statustxt(window, 'File is already compressed.')
        else:
            new_filepath = _LOCS['SCAN_DIR'] + self.name+ '.rar'
            convstr = _LOCS['WINRAR_PATH'] + ' m -m5 -ep "' + new_filepath + '" "' + self.filepath + '"'
            try:
                subprocess.call(convstr)
            except Exception as err:
                update_statustxt(window, f'Error compressing file: {sys.exc_info()[0]} - {err}')
                return False
            else:
                newbook = Book(new_filepath)
                update_statustxt(window, "File compressed successfully. New size is " + newbook.get_size_str())
                process_events.currbook = newbook
                return newbook.filename

    def format_name(self, inname):
        edfound = False
        jrfound = False
        andloc = 0

        #should deal with (ed), (ed.) and Jr/Jr.
        if '&' in inname: 
            inname = inname.replace('&', 'and')
        revname = inname.split(' ')
        if revname[-1] == '(ed)' or revname[-1] == '(ed.)':
            del revname[-1]
            edfound = True
        if revname[-1] == 'Jr' or revname[-1] == 'Jr.':
            del revname[-1]
            jrfound = True
        if 'and' in revname:
            andloc = revname.index('and')
        if 'with' in revname:
            andloc = revname.index('with')
        if 'With' in revname:
            andloc = revname.index('With')
        if 'and' not in revname or andloc == 1:
            revname.insert(0, revname.pop())
            revname = [x + '.' if len(x) == 1 else x for x in revname]
            revstr = revname[0] + ', ' + ' '.join(revname[1:])
        else:
            revstr = self.format_name(' '.join(revname[:andloc])) + ' and ' \
                     + self.format_name(' '.join(revname[andloc + 1:]))

        if edfound: 
            revstr += ' (ed.)'
        if jrfound:
            #insert 'Jr.' just before first comma signifying end of surname
            firstcomm = revstr.find(",")
            revstr = revstr[:firstcomm] + ' Jr.' + revstr[firstcomm:]
        return revstr

    def reverse_seg(self, window, segnum):
        if segnum >= len(self.seglist):
            update_statustxt(window, 'Attempt to reverse an out-of-bounds segment.')
        else:
            oldseg = self.seglist[segnum]
            self.seglist[segnum] = self.format_name(oldseg)
            self.reassemble_segs()
            return self.seglist[segnum]

    def capitalize(self):
        fixdict = {'Ii':'II', 'And ':'and ', 'In ':'in ', 'Of ':'of ', 'To ':'to ',
                   'Rtf':'rtf', 'An ':'an ', 'The ':'the ', 'Iii ':'III ', 'De ':'de ',
                   'A ':'a ', "\x92S":"'s", "'S":"'s", 'Cia ':'C.I.A. ', 'Nasa':'NASA',
                   'Kgb':'KGB', 'Mig ':'MiG ', 'Viii':'VIII', ' Iv ':' IV ', 'Fbi ':'F.B.I. ',
                   'Mcc':'McC', '(Ed.)':'(ed.)', 'Et. Al.':'et. al.', 'Trans ':'trans. ',
                   'On ':'on ', '1St':'1st', '7Th':'7th', '[Ssc]':'[SSC]',
                   'Et Al':'et. al.', 'Von ':'von ', 'Bc ':'BC ', 'Mch':'McH', 'Ss':'SS',
                   'Raf ':'R.A.F. ', "'S ":"'s ", 'Mcn':'McN', 'a. ':'A. ', ' Bc':' BC',
                   "’S":"'s", 'Wwii':'WWII', 'Mcm':'McM', 'Macn':'MacN', 'SSc':'SSC',
                   ' Iv':' IV', ' At ':' at ', '–':'-', ' By ':' by ', '40k':'40,000',
                   '40K':'40,000', 'Translated By':'translated by', 'Sf ':'SF ',
                   '2Nd':'2nd', '3Rd':'3rd', '4Th':'4th', '5Th':'5th', '6Th':'6th',
                   '8Th':'8th', '9Th':'9th', '(ed)':'(ed.)', 'IIi':'III', '10Th':'10th',
                   'Mcp':'McP', 'Gui ':'GUI ', "O'r":"O'R", 'Mcd':'McD', 'Macl':'MacL',
                   'Mcl':'McL', "n'T":"n't", ' As ':' as ', 'Ad ':'AD ', '0S':'0s', 
                   "'Ll":"'ll", 'Vs ':'vs '}

        for x in range(len(self.seglist)):
            self.seglist[x] = self.seglist[x].title()

            for k, v in fixdict.items():
                if k in self.seglist[x]: #found something that needs replacing
                    self.seglist[x] = self.seglist[x].replace(k, v)

            #capitalise first letter no matter what it is...
            if self.seglist[x] != '':
                if self.seglist[x].startswith(('[', '(')):
                    self.seglist[x] = self.seglist[x][0] + self.seglist[x][1].upper() + self.seglist[x][2:]
                else:
                    if 'translated by' not in self.seglist[x]: # ...unless it's 'translated by' string
                        self.seglist[x] = self.seglist[x][0].upper() + self.seglist[x][1:]
        self.reassemble_segs()

    def bracket_match(self, window): # clears earlier end bracket if a new, later one is added
        if self.name.count(']') > 1:
            for num, x in enumerate(self.seglist):
                if ']' in x:
                    x = x.replace(']', '')
                    self.seglist[num] = x
                    update_txtbox(window, num, x)
                    break
            self.reassemble_segs()
            display_currbook(window, False)

# ----------------------------------------------------------------------------------------

def generate_seg_layout(num):
    return [[sg.Text(f'Segment {num}:', size=(9, 1), key=f'seg{num}'), sg.Input(key=f'txt{num}',
                     enable_events=True, size=(35, 1)), sg.Button('Del', key=f'delseg{num}',
                     enable_events=True)]]

def generate_txtcols():
    return [[sg.Column(key=f'col{x}', layout = generate_seg_layout(x))] for x in range(2, NUMBOXES+1)]

def layout_window(booklist):
    sg.change_look_and_feel('DefaultNoMoreNagging')
    fontsize = 10 # note most layout is set for 11, so other sizes may misalign widgets
    buttonsize = 10
    bgcol = 'gray16'
    txtcol = 'white'
    sg.SetOptions(background_color=bgcol,
                  text_element_background_color=bgcol,
                  element_background_color=bgcol,
                  text_color=txtcol,
                  element_text_color=txtcol,
                  input_elements_background_color='gray25', #gray70
                  input_text_color='white',
                  button_color=(txtcol,'gray8'))
    #txtcols = generate_txtcols()

    col = [[sg.Text('Author(s):', size=(9, 1), pad=((10, 1),(1, 1))),
            sg.Input(key='txt1', enable_events=True, size=(30, 1)),
            sg.Button('Rev', key='btnrev'), sg.Button('Del', key='delseg1', enable_events=True)]]

    for x in generate_txtcols(): # generates rows 2 through 6 of label, input text and buttons
        col.append(x)

    sortframe = [[sg.Radio('New', 'radsort', enable_events=True, key='radnew', default=True),
                  sg.Radio('Old', 'radsort', enable_events=True, key='radold'),
                  sg.Radio('Random', 'radsort', enable_events=True, key='radrand'),
                  sg.Radio('Alpha', 'radsort', enable_events=True, key='radalpha')]]

    listcol = [[sg.Listbox(values=booklist, enable_events=True, size=(50, 8), key='filelist')],
               [sg.Frame('Sort By', sortframe),
                sg.Checkbox('Show Large (>5MB)', default=True, pad=((10, 2), (22, 2)),
                             enable_events=True, key='chklarge')]]

    layout = [[sg.Frame('Current Book', [[sg.Text('No book currently selected.', size=(86, 1), key='fullname')]]),
               sg.Frame('Size', [[sg.Text('--', size=(8, 1), justification='center', key='txtsize')]]),
               sg.Frame('Done', [[sg.Text('', size=(6, 1), justification='center', key='txtdone')]])],
              [sg.Column(listcol), sg.Column(col)],
              [sg.Text('Command:'), sg.InputText('', key='txtcmd', enable_events=True, size=(7, 1), focus=True),
               sg.Button('Go', key='btngo', pad=(10, 2), bind_return_key=True),
               sg.Text('', size=(81, 2), pad=(5, 5), background_color='gray25', key='txtstatus')],
              [sg.Button('Find Dupes', size=(buttonsize+1, 1)), sg.Button('Open', size=(buttonsize, 1)),
               sg.Button('RAR', size=(buttonsize, 1)), sg.Button('Undo', size=(buttonsize, 1)),
               sg.Button('Delete', size=(buttonsize, 1)), sg.Button('Finish', size=(buttonsize, 1)),
               sg.Button('Finish/Move', size=(11, 1)), sg.Button('Help', size=(buttonsize, 1)),
               sg.Button('Exit', size=(buttonsize, 1))]]
    window = sg.Window('Autoname 2.0', layout, margins=(3, 3), font=('Helvetica', fontsize), element_padding=(3,3),
                       icon='D:\\tmp\\renameicon.ico').Finalize()
    return window

def toggle_seg_vis(window, num, vis=True):
    # hides or unhides a selected textbox and its associated button and text
    if num: # skip author's name segment, which is index 0
        window[f'col{num+1}'].Update(visible=vis)

def update_txtbox(window, num, text=''): # these are the individual segment textboxes
    segkey = 'txt' + str(num+1)
    window[segkey].Update(text, move_cursor_to=None)
    if ']' in text:
        process_events.currbook.bracket_match(window)

def update_cmdbox(window, text=''): #command textbox
    window['txtcmd'].Update(text, move_cursor_to=None)

def update_statustxt(window, text=''): #status text
    window['txtstatus'].Update(text)

def update_done_txt(window, inc=False): # the number done textbox
    if inc: # increment the counter
        process_events.done[0] += 1
        process_events.done[1] -= 1
    window['txtdone'].Update(f'{process_events.done[0]}/{process_events.done[1]}')

def update_textboxes(window, seglist):
    # accepts a list and updates the various textboxes with it, splitting list across them
    numsegs = NUMBOXES # for now, this might be changed later

    for x in range(numsegs):
        if x < len(seglist):
            update_txtbox(window, x, seglist[x])
            toggle_seg_vis(window, x)
        else:
            update_txtbox(window, x)
            toggle_seg_vis(window, x, False)

def update_filelist(window, event, values):
    radiolist = ['radnew', 'radrand', 'radold', 'radalpha']
    showlarge = window['chklarge'].Get()

    if values and event not in radiolist:
        #different event triggered filelist update, like a rename, so get current sort setting
        for x in radiolist:
            if values[x]:
                event = x
                break

    if event == 'radnew':
        booklist = gen_booklist('newestfirst', showlarge)
    elif event == 'radrand':
        booklist = gen_booklist('random', showlarge)
    elif event == 'radold':
        booklist = gen_booklist('oldestfirst', showlarge)
    elif event == 'radalpha':
        booklist = gen_booklist('alphabetical', showlarge)

    window['filelist'].Update(values=booklist, set_to_index=0)
    ab = window['filelist'].GetListValues()
    if ab != []:
        process_events.currbook = Book(_LOCS['SCAN_DIR'] + ab[0])
        display_currbook(window)
        process_events.done[1] = len(booklist)
        update_done_txt(window)
        return booklist
    else:
        return None

def show_help():
    # ddd, u(ndo), f, fff, ca, cd, delseg, [X, ]X, 40k  '\n• '
    helptext = 'Available Commands:' \
    '\n• rX: Reverse segment X, defaults to first segment.' \
    '\n• c: Capitalize book title.' \
    '\n• as: Add a segment.' \
    '\n• XY: Swap positions of segments X and Y.' \
    '\n• splX: split segment X, defaults to first segment.' \
    '\n• by: Remove "by" and split segment.' \
    '\n• o: Open current file.' \
    '\n• rar: Compress current file.' \
    '\n• fd: Find duplicates of file in target directory.' \
    "\n• f: Finalize book name but don't move." \
    '\n• fff: Finalize book name and move to output directory.' \
    '\n• [X, X]: Add square brackets to start or end of segment X.' \
    '\n• ddd: Delete book after confirmation.' \
    "\n• trans: Add 'translated by ' to end of title" \
    "\n• ed: Add '(ed.)' to end of author's name" \
    '\n• ssc, 40k: Add identifier tags to book title' \
    '\n• d-X: Deletes all hyphens from segment X, defaults to first.' \
    '\n• d[X: Deletes all square brackets [] from segment X, defaults to first.' \
    '\n• d[X: Deletes all brackets () from segment X, defaults to first.' \
    '\n• d.X: Delete all full stops from segment X, defaults to first.' \
    '\n• d_X: Delete all underscores from segment X, defaults to first.' \
    '\n• q: Quit.'
    sg.PopupOK(helptext, title='Help')

def gen_booklist(mode='newestfirst', showlarge=True):
    fdict = {}
    filelist = []

    for book in glob(_LOCS['SCAN_DIR'] + '*.rar') + glob(_LOCS['SCAN_DIR'] + '*.pdf') +\
                glob(_LOCS['SCAN_DIR'] + '*.txt'):
        if showlarge == False:
            size =  os.path.getsize(book)
            if size > 5000000:
                continue
        fdict[book] = os.path.getmtime(book)

    for book, date in sorted(iter(fdict.items()), key=lambda x: x[1], reverse=(mode=='newestfirst')):
        filelist.append(os.path.basename(book))

    if mode == 'random':
        shuffle(filelist)
    if mode == 'alphabetical':
        filelist = sorted(filelist)

    return filelist

def open_bookfile(window):
    #open file up to take a look inside
    try:
        os.startfile(process_events.currbook.filepath)
    except Exception as err:
        update_statustxt(window, f'Error opening file - {err}')

def move_to_next_book(window, lastbook=None):
    #lastbook is 'delete', 'revert' or 'retain'
    #Finish/Move uses delete, Finish uses retain, moving onwards normally uses revert
    allbooks = window['filelist'].GetListValues()
    if lastbook == 'delete': # delete old book entry and reload list, then highlist next book
        del allbooks[process_events.currindex]
    elif lastbook == 'retain':
        allbooks[process_events.currindex] = process_events.currbook.filename
        process_events.currindex += 1
    elif lastbook == 'revert':
         #if process_events.currbook.name != allbooks[process_events.currindex]:
            #allbooks
        process_events.currindex += 1

    window['filelist'].Update(values=allbooks, set_to_index=process_events.currindex,
                              scroll_to_index=process_events.currindex)
    if process_events.currindex >= len(allbooks):
        process_events.currindex = len(allbooks) - 1
    if allbooks:
        newbook = allbooks[process_events.currindex]
        process_events.currbook = Book(_LOCS['SCAN_DIR'] + newbook)
    else: #empty list as no books are left
        process_events.currindex = 0
        process_events.currbook = None

def move_to_specified_book(window, bookname):
    allbooks = window['filelist'].GetListValues()
    bookpos = allbooks.index(bookname)
    process_events.currindex = bookpos
    window['filelist'].Update(set_to_index=bookpos, scroll_to_index=bookpos)
    process_events.currbook = Book(_LOCS['SCAN_DIR'] + bookname)

def process_txt_cmd(window, values, cmd):
    update_cmdbox(window)
    if cmd == 'q':
        #print('Quittng.')
        window.Close()
        sys.exit()
    elif cmd == '': #move to next book
        #move_to_next_book(window, 'revert') # this was fine as a text app, but as a gui
        pass                                 # this causes more problems than it's worth
    elif cmd == 'fd':
        process_events.currbook.dupefinder(window)
    elif cmd == 'by':
        process_events.currbook.by_replace()
    elif cmd == 'rar':
        newbook = process_events.currbook.rar(window)
        if newbook: #update listbox filename too
            # need to get current listbox value, rename to .rar, and then update
            update_filelist(window, None, values)
            move_to_specified_book(window, newbook)
    elif cmd[0] == 'r':  # reverse
        if len(cmd) == 1:
            cmd += '1'  #allow just 'r' to reverse segment 1
        segnum = int(cmd[1]) - 1
        oldseg = process_events.currbook.seglist[segnum]
        revseg = process_events.currbook.reverse_seg(window, segnum)
        update_statustxt(window, f'Reversing {oldseg} to {revseg}.')
    elif cmd == 'o':
        open_bookfile(window)
    elif cmd == 'h':
        show_help()
    elif cmd[0] == 'c': #capitalise segment and fix short words
        #if len(cmd) == 1: cmd = 'c1'
        process_events.currbook.capitalize()
    elif cmd[:3] == 'spl': #split a segment:
        segnum = int(cmd[3]) - 1 if len(cmd) == 4 else 0  # segnum is 0-based not 1
        process_events.currbook.split_seg(window, segnum)
    elif len(cmd) == 2 and cmd.isdigit(): # two numbers, swap segments
        process_events.currbook.swap_segs(window, int(cmd[0]), int(cmd[1]))
    elif cmd == 'fff': #rename and move to output dir
        res = process_events.currbook.finish(window, True)
        if res:
            move_to_next_book(window, 'delete')
            update_done_txt(window, True)
    elif cmd == 'f': # rename book but don't move
        res = process_events.currbook.finish(window, False)
        if res:
            move_to_next_book(window, 'retain')
    elif cmd == 'as': # add a new segment
        process_events.currbook.add_seg()
    elif cmd == '40k': # Ave Imperator!
        process_events.currbook.add_seg('[Warhammer 40,000', 1)
    elif cmd == 'ssc': #short story collection designator
        process_events.currbook.add_seg('[SSC]', 1)
    elif cmd == 'trans': # add 'translated by'
        process_events.currbook.add_seg('translated by ')
    elif cmd == 'ed': # add (ed.) to end of author's name
        process_events.currbook.seglist[0] = process_events.currbook.seglist[0] + ' (ed.)'
        process_events.currbook.reassemble_segs()
    elif cmd[0] == '[': # square bracket to beginning of segment
        segnum = int(cmd[1]) if len(cmd) > 1 else 0
        process_events.currbook.seglist[segnum-1] = '[' + process_events.currbook.seglist[segnum-1]
        process_events.currbook.reassemble_segs()
    elif cmd[0] == ']': # square bracket to end of segment
        segnum = int(cmd[1]) if len(cmd) > 1 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1] + ']'
        process_events.currbook.reassemble_segs()
    elif cmd == 'undo': # revert all changes
        origname = process_events.currbook.filepath
        process_events.currbook = Book(origname)
    elif cmd == 'ddd': # delete current file
        res = process_events.currbook.delete(window)
        if res:
            move_to_next_book(window, 'delete')
    elif cmd[:2] == 'd-': # delete all hyphens from segment
        segnum = int(cmd[2]) if len(cmd) > 2 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace('-',' ')
        process_events.currbook.reassemble_segs()
    elif cmd[:2] == 'd[': # delete all square brackets from segment
        segnum = int(cmd[2]) if len(cmd) > 2 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace('[','')
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace(']','')
        process_events.currbook.reassemble_segs()
    elif cmd[:2] == 'd.': # delete all full stops from segment
        segnum = int(cmd[2]) if len(cmd) > 2 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace('.',' ')
        process_events.currbook.reassemble_segs()
    elif cmd[:2] == 'd_': # delete all underscores from segment
        segnum = int(cmd[2]) if len(cmd) > 2 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace('_',' ')
        process_events.currbook.reassemble_segs()
    elif cmd[:2] == 'd(': # delete all brackets from segment
        segnum = int(cmd[2]) if len(cmd) > 2 else 0
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace('(',' ')
        process_events.currbook.seglist[segnum-1] = process_events.currbook.seglist[segnum-1].replace(')',' ')
        process_events.currbook.reassemble_segs()
    else:
        update_statustxt(window, 'Command not recognised.')

def process_events(window, event, values):
    txtboxes = ['txt' + str(x) for x in range(1, NUMBOXES+1)] # doing this every time doesn't seem efficient
    acceptletts = string.ascii_letters + string.digits + " []()-&,.;'"

    if not process_events.currbook and event != 'Help': # no books so disable all buttons except Help
        return

    if event == 'filelist': #update the file list window
        listedbookname = values['filelist'][0]
        process_events.currbook = Book(_LOCS['SCAN_DIR'] + listedbookname)
        process_events.currindex = window['filelist'].Widget.curselection()[0]
    elif event == 'btngo': # a text command is to be executed
        cmd = window['txtcmd'].Get()
        try:
            process_txt_cmd(window, values, cmd)
        except Exception as err:
            update_statustxt(window, f'Command \'{cmd}\' not recognised or invalid - {err}')
    elif event == 'btnrev':
        oldseg = process_events.currbook.seglist[0]
        revseg = process_events.currbook.reverse_seg(window, 0)
        update_statustxt(window, f'Reversing {oldseg} to {revseg}.')
    elif event == 'RAR':
        #oldname = values['filelist']
        newbook = process_events.currbook.rar(window)
        if newbook: #update listbox filename too
            # need to get current listbox value, rename to .rar, and then update
            update_filelist(window, event, values)
            move_to_specified_book(window, newbook)
    elif event == 'Find Dupes':
        process_events.currbook.dupefinder(window)
    elif event == 'Open':
        open_bookfile(window)
    elif event == 'Help':
        show_help()
    elif event == 'Undo':
        origname = process_events.currbook.filepath
        process_events.currbook = Book(origname)
    elif event == 'Finish/Move':
        res = process_events.currbook.finish(window, True)
        if res:
            move_to_next_book(window, 'delete')
            update_done_txt(window, True)
    elif event == 'Finish':
        res = process_events.currbook.finish(window, False)
        if res:
            move_to_next_book(window, 'retain')
    elif event == 'Delete':
        res = process_events.currbook.delete(window)
        if res:
            move_to_next_book(window, 'delete')
    elif event == 'chklarge':
        update_filelist(window, event, values)
    elif 'delseg' in event: # one of the individual delete segment buttons
        getsegnum = int(event[-1])
        process_events.currbook.del_seg(window, getsegnum)
    elif event in ['radold', 'radnew', 'radalpha', 'radrand']: # file list sort options
        update_filelist(window, event, values)
    elif event in txtboxes: # rebuild filename with edited text
        for num, x in enumerate(process_events.currbook.seglist):
            key = 'txt'+str(num+1)
            if key in values:
                txtboxdata = values[key]
                #validate data
                txtboxdata = ''.join([x for x in txtboxdata if x in acceptletts])
                if txtboxdata != process_events.currbook.seglist[num]:
                    process_events.currbook.edit_seg(num, txtboxdata)
            else:
                update_statustxt(window, 'Error renaming book: invalid key.')

    # if editing a text box, don't want focus to snap back to cmd txtbox
    display_currbook(window, values, event not in txtboxes)

def display_currbook(window, values=None, resetfocus=True):
    #takes care of displaying current book's details at the top
    currbook = process_events.currbook
    if currbook:
        window['fullname'].Update(f'{currbook.filename}') # ({currbook.size})

        booksize = currbook.get_size_int()
        if booksize:
            txtcol = 'white' if booksize < 5000 else 'red'
            window['txtsize'].Update(f'{currbook.size}', text_color=txtcol)
        else:
            update_statustxt(window, 'Selected book has been moved, deleted or renamed.'\
                                     ' Refreshing file list.')
            update_filelist(window, None, values)

        update_textboxes(window, currbook.seglist)
        if resetfocus:
            window['txtcmd'].SetFocus()
    else:
        update_statustxt(window, f'No books found in current working directory ({_LOCS["SCAN_DIR"]}).')

def load_config():
    config = ConfigParser()
    try:
        config.read('psg-autoname.ini')
        _LOCS['SCAN_DIR'] = config['Locations']['scandir']
        _LOCS['OUTPUT_DIR'] = config['Locations']['outputdir']
        _LOCS['WINRAR_PATH'] = config['Locations']['winrarpath']
        return True
    except Exception as err:
        print(f'Error reading or parsing config file - {err}')
        return False

def dir_loader():
    load_dir = os.listdir(_LOCS['OUTPUT_DIR'])

def start_preloader():
    '''a major problem has been that the dupefinder function hangs for ~30 secs because
    of the slowness of scanning a dir with ~20K files on an old laptop. This function spins 
    off a thread to do that and thus force the OS to cache results on program startup, speeding 
    all future uses of the dupefinder'''
    t = threading.Thread(target=dir_loader)
    t.start()

def main():
    if not load_config():
        sys.exit(1)
    start_preloader()
    booklist = gen_booklist()
    if booklist:
        currbook = Book(_LOCS['SCAN_DIR'] + booklist[0])
        process_events.currbook = currbook
        process_events.currindex = 0
    else:
        currbook = None
        process_events.currbook = None
    process_events.done = [0, len(booklist)]
    window = layout_window(booklist)
    window['txtdone'].Update(value=f'0/{process_events.done[1]}')
    display_currbook(window)

    while True:
        event, values = window.Read()
        #print(event, values)
        if event is None or event == 'Exit':
            break
        elif event == '__TIMEOUT__':
            pass
        else:
            process_events(window, event, values)

    window.Close()

if __name__ == '__main__':
    main()
