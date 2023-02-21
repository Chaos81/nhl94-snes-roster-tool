# """ Program extracts roster data from SNES NHL '94 ROMs."""
# """ Version 0.7.0 """
# Version history
# 0.1.0 - Original PHP script
# 0.2.0 - Original Python version, command line interface
# 0.3.0 - GUI Version with File Dialog and Buttons, Export only
# 0.4.0 - Bug fixes, Help menu added
# 0.5.0 - Import to CSV added
# 0.5.5 - Added Overall Rating
# 0.6.0 - Bug fix (Windows CSV Output), SMC Header check
# 0.6.5 - Bug fix (Finish process of importing of last team, Player Stats bug - add 0200 to end of roster list)
# 0.7.0 - Added reading the Player Data Offset Bytes

from tkinter import Tk, Menu, PhotoImage, BOTH
from tkinter.ttk import Frame, Button, Label
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showinfo, showerror

import sys
from binascii import b2a_hex
import csv
import os
import shutil
import struct


class RosExt(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)

        self.parent = parent

        # Instance Variables
        self.bg_image = ""
        self.head_offset = 0  # Header Offset

        self.initUI()



    def initUI(self):

        self.parent.title("SNES NHL '94 Roster Tool Ver. 0.70")
        self.pack(fill=BOTH, expand=1)

        # Menu
        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)

        fileMenu = Menu(menubar, tearoff=0)
        fileMenu.add_command(label="Import ROM to CSV...", command=self.importcsv)
        fileMenu.add_command(label="Export ROM to CSV...", command=self.extractrom)
        fileMenu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=fileMenu)

        helpMenu = Menu(menubar, tearoff=0)
        helpMenu.add_command(label="Export to CSV Instructions...", command=self.expinst)
        helpMenu.add_command(label="Import from CSV Instructions...", command=self.impinst)

        helpMenu.add_command(label="About...", command=self.about)
        menubar.add_cascade(label="Help", menu=helpMenu)

        # Background Image
        image = self.find_data_file('nhl94.gif')
        self.bg_image = PhotoImage(file=image)
        bg_label = Label(self, image=self.bg_image)
        bg_label.grid(row=0, column=0, padx=30, pady=20, columnspan=2)
        # Buttons
        export_button = Button(self, text="Export to CSV", command=self.extractrom)
        export_button.grid(row=1, column=1)
        import_button = Button(self, text="Import from CSV", command=self.importcsv)
        import_button.grid(row=1, column=0)

    def find_data_file(self, filename):
        if getattr(sys, "frozen", False):
            # The application is frozen.
            datadir = os.path.dirname(sys.executable)
        else:
            # The application is not frozen.
            datadir = os.path.dirname(__file__)

        return os.path.join(datadir, filename)

    def expinst(self):

        showinfo("SNES Roster Tool Instructions",
                 "The Roster tool allows you to export a ROM's roster data to a CSV file.\n"
                 "CSV format is a common format used in spreadsheet programs like MS Excel.\n\n"
                 "Exported data can easily be sorted in a spreadsheet program.  The exported data includes the name of "
                 "the player, the team they are playing for, and all of the player's attributes (Weight, ShP, etc). "
                 "These stats are displayed on a scale of 0-6 (Weight and Handedness are on a scale of 0-15).\n\n"
                 "Stats are:\n 0 = 25, 1 = 35, 2 = 45, 3 = 54, 4 = 72, 5 = 90, 6 = 100 \n\n"
                 "Weights are:\n 0 = 140, 1 = 148, 2 = 156, 3 = 164, 4 = 172, 5 = 180, 6 = 188, 7 = 196, 8 = 204 "
                 "9 = 212, 10 = 220, 11 = 228, 12 = 236, 13 = 244, 14 = 252, 15 = 260\n\n"
                 "Handedness is an odd number for R, even number for L\n\n"
                 "The program will ask you to choose the NHL '94 ROM file to extract the roster data from.  Please "
                 "choose the ROM file, then it will ask you to choose a filename to save the data to in CSV format."
                 "  The program will tell you if it completed successfully, or if there was an error.")

    def impinst(self):

        showinfo("SNES Roster Tool Instructions",
                 "The Roster tool allows you to import roster data into a NHL '94 ROM from a CSV formatted file.\n"
                 "This will allow you to make widespread changes to rosters easily.\n"
                 "CSV format is a common format used in spreadsheet programs like MS Excel.\n\n"
                 "Imported data needs to be in a specific order in order for it to import successfully.\n"
                 "The program expects the data to be in the following columns:\n"
                 "First, Last, Abv, Pos, JNo, Ovr, Wgt, Agl, Spd, OfA, DfA, Shp-PkC, Chk, Hnd, StH, ShA, End-StR, "
                 "Rgh-StL, "
                 "Pas-GlR, Agr-GlL\n\n"
                 "The program will detect if the first row of the file contains the column labels listed above.  "
                 "If you use column labels, the data can be in any order, but if you choose not use column labels, the "
                 "data needs to be stored in the above order.\n\nYou can use the CSV file output from 'Export to CSV' "
                 "as a template.\n\n"
                 "The program will ask for the CSV file to import, then will ask you for the ROM file you wish to "
                 "import to.  It will then ask you to choose a name and location to save the modified ROM.  It will "
                 "make a copy of the ROM file, import the rosters, and save the modified copy to the named location.  "
                 "The program will tell you if it completed successfully, or if there was an error and what is needed "
                 "to correct it.\n\n"
                 "NOTE: The player data space is limited to the same size as the original NHL '94 ROM rosters.  The "
                 "program will notify you when there is no more space in the ROM for that team.\n\nThe ROM file created"
                 " with this program is compatible with the SNES Editor created by Statto.")

    def about(self):

        showinfo("About SNES Roster Tool", "SNES Roster Tool Version 0.7\n\nCreated by chaos\n\nIf there are any bugs "
                                           "or questions, please email me at chaos@nhl94.com")

    def checkhead(self, f):

        # Checks for SMC header and creates offset if needed
        # Checks for ROM Name in ROM Header at 32704 (7FC0) - NHL '94 (4E 48 4C 20 27 39 34)
        # Header is size 512 bytes (200 hex)

        f.seek(32704)
        name = b2a_hex(f.read(7)).decode("utf-8")

        print(name)

        if name == "4e484c20273934":
            self.head_offset = 0
        else:
            self.head_offset = 512

        print(self.head_offset)

    def importcsv(self):

        ftypes = [("'CSV Files", '*.csv')]
        romtypes = [("'94 ROM Files", '*.smc')]
        home = os.path.expanduser('~')
        csvfile = askopenfilename(title="Choose a CSV file for import...", filetypes=ftypes, initialdir=home)
        if csvfile != '':
            try:
                rom = askopenfilename(title="Choose a '94 ROM file...", filetypes=romtypes, initialdir=home)
                save = asksaveasfilename(title="Enter a name for the new '94 ROM file...",
                                         defaultextension='.smc', initialdir=home)
                shutil.copyfile(rom, save)

                with open(save, 'rb+') as f:
                    success = self.importroster(csvfile, f)
                    f.close()
                    if success == 0:
                        showinfo("SNES NHL '94 Roster Tool", "Roster Data has been imported successfully.")
                    elif success == 1:
                        showerror("SNES NHL '94 Roster Tool", "There is an error.")
                    elif success == 2:
                        showerror("SNES NHL '94 Roster Tool", "The CSV file is missing fields or some fields are "
                                                              "blank.  Please check the file.")
                    elif success == 3:
                        showerror("SNES NHL '94 Roster Tool", "The CSV file has a team listed that cannot be found in "
                                                              "the ROM.  Please check the file.")
                    elif success == 4:
                        showerror("SNES NHL '94 Roster Tool", "Please make the necessary changes to the CSV file and "
                                                              "try again.")

            except IOError:
                showerror("SNES NHL '94 Roster Tool", "Could not open ROM or CSV file.  Please check file "
                                                      "permissions.")
            except ValueError:
                showerror("SNES NHL '94 Roster Tool",
                          "There was an error in accessing or modifying roster info.  Please make sure that "
                          "you are using a valid NHL '94 ROM and the CSV file is formatted correctly.")

    def extractrom(self):

        ftypes = [("'94 ROM Files", '*.smc')]
        home = os.path.expanduser('~')
        file = askopenfilename(title="Please choose a '94 ROM file...", filetypes=ftypes, initialdir=home)
        if file != '':
            try:
                f = open(file, 'rb')
                save = asksaveasfilename(title="Please choose a name and location for the CSV file...",
                                         defaultextension='.csv', initialdir=home)
                with open(save, 'w', newline='') as w:
                    self.extractroster(f, w)
                    f.close()
                    w.close()
            except IOError:
                showerror("SNES NHL '94 Roster Tool", "Could not open ROM or create CSV file.  Please check file "
                                                      "permissions.")
            except ValueError:
                showerror("SNES NHL '94 Roster Tool",
                          "There was an error in retreiving roster info.  Please make sure that "
                          "you are using a valid NHL '94 ROM.")

    def importroster(self, file, f):
        # Import roster data from CSV into ROM

        fields = ['First', 'Last', 'Abv', 'Pos', 'JNo', 'Ovr', 'Wgt', 'Agl', 'Spd', 'OfA', 'DfA', 'ShP-PkC', 'Chk',
                  'Hnd', 'StH', 'ShA', 'End-StR', 'Rgh-StL', 'Pas-GlR', 'Agr-GlL']
        tminfo = {}
        tmdata = {}

        # CSV open

        csvfile = open(file, 'r', newline='')

        # Check for Header Rows

        header = csv.Sniffer().has_header(csvfile.read(1024))
        csvfile.seek(0)

        if header:
            reader = csv.DictReader(csvfile)
            chk = set(fields) & set(reader.fieldnames)
            if len(chk) != 20:
                showerror("SNES '94 Roster Tool",
                          "The column field names are incorrect.  They should be: " + ', '.join(fields))
                csvfile.close()
                return 4

        else:
            reader = csv.DictReader(csvfile, fieldnames=fields)

        check = self.check_csv(reader)
        if not check:
            csvfile.close()
            return 2

        # If Header row, it must be skipped

        if header:
            csvfile.seek(0)
            next(csvfile)
        else:
            csvfile.seek(0)

        # Retrieve Team Pointers and Info from ROM
        tmarray = self.tm_ptrs(f)

        # Generate Dictionary containing pointer, player space, and player data offset for each team
        for ptr in tmarray:
            tminfo = self.get_team_info(f, ptr)
            tmlist = [ptr, tminfo['plspace'], tminfo['ploff']]
            tmdata[tminfo['abv']] = tmlist

        # Compare CSV Data Size to Player Data Size
        # Each Player has a set amount of bytes, with only "Name" as variable.  Bytes = 10 + Player Name

        curtm = ""
        count = 0
        numg = numf = numd = 0
        tmspace = 0
        tmptr = 0

        for row in reader:
            if row['Abv'] != curtm:
                prevtm = curtm
                curtm = row['Abv']
                # Check and see if team is in ROM
                if curtm in tmdata:
                    if count != 0:  # Not the First Run, change Number of Players
                        # Bug Fix - Add 02 00 to end of Roster List
                        f.write(struct.pack("B", 2))
                        f.write(struct.pack("B", 0))
                        # Pad rest of Player data with FF
                        for num in range(0, tmspace - 2):  # -2 to compensate for the 0200 above
                            f.write(struct.pack("B", 255))

                        # Prepare and write previous team's G, F and D
                        if numg == 1:
                            gb1 = 16
                            gb2 = 0
                        elif numg == 2:
                            gb1 = 17
                            gb2 = 0
                        elif numg == 3:
                            gb1 = 17
                            gb2 = 16
                        elif numg == 4:
                            gb1 = 17
                            gb2 = 17
                        else:
                            showerror("SNES '94 Roster Tool",
                                      "The number of Goalies on " + prevtm + " must be between 1 and 4.")
                            csvfile.close()
                            return 4

                        if numf < 1 or numf > 15:
                            showerror("SNES '94 Roster Tool",
                                      "The number of Forwards on " + prevtm + " must be between 1 and 15.")
                            csvfile.close()
                            return 4
                        if numd < 1 or numd > 15:
                            showerror("SNES '94 Roster Tool",
                                      "The number of Defenders on " + prevtm + " must be between 1 and 15.")
                            csvfile.close()
                            return 4

                        fwd = hex(numf)
                        fwd = fwd[2:]
                        dfs = hex(numd)
                        dfs = dfs[2:]
                        numfd = fwd + dfs

                        # Update G, F and D

                        f.seek(tmptr + 17)
                        f.write(struct.pack("B", int(numfd, 16)))
                        f.seek(tmptr + 19)
                        f.write(struct.pack("2B", gb1, gb2))

                        # Update Lines (First G, First 3 Fs, First 2 D)
                        # BEST, SC1, SC2, CHK, PP1, PP2, PK1, PK2 - G, LD, RD, LW, C, RW, XA

                        # Find Players for Default Line
                        line = [hex(1)]  # G
                        firstd = numg + numf + 1
                        firstf = numg + 1
                        line.append(hex(firstd))
                        line.append(hex(firstd + 1))
                        for fwd in range(firstf, firstf + 4):
                            line.append(hex(fwd))

                        f.seek(tmptr + 21)  # Move to BEST

                        for i in range(1, 9):
                            for plr in line:
                                f.write(struct.pack("B", int(plr, 16)))
                            f.write(struct.pack("B", 0))

                        numg = numf = numd = 0

                    count += 1
                    tmptr = tmdata[curtm][0]
                    tmspace = tmdata[curtm][1]
                    tmspace = int(tmspace)
                    f.seek(tmptr + tmdata[curtm][2])  # Move to Player Data
                else:
                    csvfile.close()
                    return 3

            # Process row from CSV
            # Name passed to bytearray for conversion and writing in ASCII format.  All other values are converted to
            # int base 16 before being written into file

            name = row['First'] + " " + row['Last']
            nmlength = len(name) + 2  # First Byte

            # Check to see if there is room for Player

            if tmspace <= 10:
                showerror("SNES '94 Roster Tool", "There is not enough player space to add " + name + " for team "
                          + curtm + ".  You are " + str(abs(tmspace)) + " player bytes short.")
                csvfile.close()
                return 4

            tmspace -= nmlength + 8

            if row['Pos'] == 'G':
                numg += 1
            elif row['Pos'] == 'F':
                numf += 1
            elif row['Pos'] == 'D':
                numd += 1
            else:
                showerror("SNES '94 Roster Tool", "There is an unknown position designated for " + name + ".")
                csvfile.close()
                return 4

            # Begin writing to file

            f.write(struct.pack("B", nmlength))
            f.write(struct.pack("B", 0))
            player = bytearray(name, 'utf-8')
            f.write(player)

            jno = row['JNo']
            f.write(struct.pack("B", int(jno, 16)))

            # Change Weight and Handedness to Hex

            wgt = hex(int(row['Wgt']))
            wgt = wgt[2:]

            hnd = hex(int(row['Hnd']))
            hnd = hnd[2:]

            # Create Attrib String

            attstring = wgt + row['Agl'] + row['Spd'] + row['OfA'] + row['DfA'] + row['ShP-PkC'] + row['Chk']
            attstring += hnd + row['StH'] + row['ShA'] + row['End-StR'] + row['Rgh-StL'] + row['Pas-GlR']
            attstring += row['Agr-GlL']
            attrib = [attstring[i:i + 2] for i in range(0, len(attstring), 2)]
            for att in attrib:
                f.write(struct.pack("B", int(att, 16)))

        # Process the last team in CSV file
        if count != 0:  # At least 1 team was processed
            # Pad rest of Player data with FF
            # Bug Fix - Add 02 00 to end of Roster List
            f.write(struct.pack("B", 2))
            f.write(struct.pack("B", 0))
            for num in range(0, tmspace - 2):  # -2 to compensate for the 0200 above
                f.write(struct.pack("B", 255))

            # Prepare and write previous team's G, F and D
            if numg == 1:
                gb1 = 16
                gb2 = 0
            elif numg == 2:
                gb1 = 17
                gb2 = 0
            elif numg == 3:
                gb1 = 17
                gb2 = 16
            elif numg == 4:
                gb1 = 17
                gb2 = 17
            else:
                showerror("SNES '94 Roster Tool",
                          "The number of Goalies on " + prevtm + " must be between 1 and 4.")
                csvfile.close()
                return 4

            if numf < 1 or numf > 15:
                showerror("SNES '94 Roster Tool",
                          "The number of Forwards on " + prevtm + " must be between 1 and 15.")
                csvfile.close()
                return 4
            if numd < 1 or numd > 15:
                showerror("SNES '94 Roster Tool",
                          "The number of Defenders on " + prevtm + " must be between 1 and 15.")
                csvfile.close()
                return 4

            fwd = hex(numf)
            fwd = fwd[2:]
            dfs = hex(numd)
            dfs = dfs[2:]
            numfd = fwd + dfs

            # Update G, F and D

            f.seek(tmptr + 17)
            f.write(struct.pack("B", int(numfd, 16)))
            f.seek(tmptr + 19)
            f.write(struct.pack("2B", gb1, gb2))

            # Update Lines (First G, First 3 Fs, First 2 D)
            # BEST, SC1, SC2, CHK, PP1, PP2, PK1, PK2 - G, LD, RD, LW, C, RW, XA

            # Find Players for Default Line
            line = [hex(1)]  # G
            firstd = numg + numf + 1
            firstf = numg + 1
            line.append(hex(firstd))
            line.append(hex(firstd + 1))
            for fwd in range(firstf, firstf + 4):
                line.append(hex(fwd))

            f.seek(tmptr + 21)  # Move to BEST

            for i in range(1, 9):
                for plr in line:
                    f.write(struct.pack("B", int(plr, 16)))
                f.write(struct.pack("B", 0))

        return 0

    def extractroster(self, f, w):
        # Extract roster data from ROM

        fields = ['First', 'Last', 'Abv', 'Pos', 'JNo', 'Ovr', 'Wgt', 'Agl', 'Spd', 'OfA', 'DfA', 'ShP-PkC', 'Chk',
                  'Hnd', 'StH', 'ShA', 'End-StR', 'Rgh-StL', 'Pas-GlR', 'Agr-GlL']
        tminfo = {}
        writer = csv.DictWriter(w, fieldnames=fields, delimiter=',')
        writer.writeheader()
        tmarray = self.tm_ptrs(f)
        print(tmarray)

        for ptr in tmarray:
            tminfo = self.get_team_info(f, ptr)
            self.get_player_info(f, ptr, tminfo, writer)

        showinfo("SNES NHL '94 Roster Tool", "Roster Data has been extracted.")

    def check_csv(self, reader):
        # Check CSV to make sure there are no missing fields for each entry

        for row in reader:
            if any(val in ("") for val in row.values()):
                return False
        return 1

    def lit_to_big(self, little):
        # Change byte string from little to big endian
        return little[2:4] + little[0:2]

    def tm_ptrs(self, f):
        # Retrieve Team Offset Pointers

        # Check for Header
        self.checkhead(f)

        # Team Offset Start Position - 927207 - Headerless, 927719 Headered
        f.seek(927207 + self.head_offset)

        ptrarray = []

        for i in range(0, 28):
            firsttm = b2a_hex(f.read(2))
            f.seek(2, 1)

            conv = self.lit_to_big(firsttm)

            # If needed, add header offset
            if self.head_offset == 512:
                data = int(conv, 16) + int('0x0D8200', 16)
            else:
                data = int(conv, 16) + int('0x0D8000', 16)

            print(data)
            ptrarray.append(data)

        return ptrarray

    def get_team_info(self, f, ptr):
        # Retrieve Team Info

        # Team Name Data starts at the end of Player Data (offset given at bytes 4 and 5 in Team Data)
        # First offset: Length of Team City (including this byte)
        # AA 00 TEAM CITY BB 00 TEAM ABV CC 00 TEAM NICKNAME DD 00 TEAM ARENA
        # AA - Length of Team City (includes AA and 00)
        # BB - Length of Team Abv (includes BB and 00)
        # CC - Length of Team Nickname (includes CC and 00)
        # DD - Length of Team Arena (includes DD and 00)
        # All Name Data is in ASCII format.

        # Player Data Offset - Default is 55 00 (85 bytes), but in some custom ROMs, may be different
        f.seek(ptr)
        plpos = self.lit_to_big(b2a_hex(f.read(2)))
        ploff = int(plpos, 16)

        # Team Data Offset - Team Offset + 4 bytes
        f.seek(ptr + 4)
        tmpos = self.lit_to_big(b2a_hex(f.read(2)))
        dataoff = ptr + int(tmpos, 16)

        # Calculate Player Data Space
        # Team Data Offset - Player Data Offset - 2 (last 2 bytes of Player Data 02 00)
        plsize = int(tmpos, 16) - ploff - 2

        # Read Team City
        f.seek(dataoff)
        tml = int(self.lit_to_big(b2a_hex(f.read(1))), 16)
        # Skip 00
        f.seek(1, 1)
        tmcity = f.read(tml - 2).decode("utf-8")

        # Read Team Abv
        tml = int(self.lit_to_big(b2a_hex(f.read(1))), 16)
        f.seek(1, 1)
        tmabv = f.read(tml - 2).decode("utf-8")

        # Read Team Nickname
        tml = int(self.lit_to_big(b2a_hex(f.read(1))), 16)
        f.seek(1, 1)
        tmnm = f.read(tml - 2).decode("utf-8")

        return dict(city=tmcity, abv=tmabv, name=tmnm, plspace=plsize, ploff=ploff)

    def get_player_info(self, f, ptr, tminfo, writer):
        # Retreive Player Info

        # Player Data Starts 85 bytes (0x55) from Start offset (may be different in custom ROM)

        # XX 00 "PLAYER NAME" XX 123456789ABCDE

        # XX =	"Player name length" + 2 (the two bytes in front of the name) in hex.
        # 00 =	Null (Nothing)

        # "PLAYER NAME"

        # XX =	Jersey # (decimal)

        # 1 = Weight
        # 2 = Agility

        # 3 = Speed
        # 4 = Off. Aware.

        # 5 = Def. Aware.
        # 6 = Shot Power/Puck Control

        # 7 = Checking
        # 8 = Stick Hand (Uneven = Right. Even = Left. 0/1 will do.)

        # 9 = Stick Handling
        # A = Shot Accuracy

        # B = Endurance/StR
        # C = ? (Roughness on Genesis)/StL

        # D = Passing/GlR
        # E = Aggression/GlL

        # Calculate # of Players - Goalies First, then F and D

        f.seek(ptr + 19)
        gdata = b2a_hex(f.read(2)).decode("utf-8")
        numg = gdata.find("0")
        f.seek(ptr + 17)

        pdata = b2a_hex(f.read(1))
        numf = int(pdata[0:1], 16)
        numd = int(pdata[1:2], 16)

        nump = numg + numf + numd

        # Move to Player Data

        f.seek(ptr + tminfo['ploff'])
        for i in range(1, nump + 1):
            # Name and JNo
            pnl = int(b2a_hex(f.read(1)), 16)
            f.seek(1, 1)
            name = f.read(pnl - 2).decode("utf-8")
            jno = b2a_hex(f.read(1)).decode("utf-8")
            names = name.split(" ")

            # G, F or D?

            if i <= numg:
                pos = 'G'
            elif i <= (numg + numf):
                pos = 'F'
            else:
                pos = 'D'

            # Get Attributes

            attrib = []
            adata = b2a_hex(f.read(7)).decode("utf-8")
            for ch in adata:
                stat = int(ch, 16)
                attrib.append(stat)

            # Calculate Overall Ratings

            # PLAYER:
            # total = (agility * 2) + (speed * 3) + (offensive * 3) + (defensive * 2) + (shot_power * 1)
            # + (checking * 2) + (stick_handling_value * 3) + (shot_accuracy * 2) + (endurance * 1) + (pass * 1)

            # if total < 50
            #    x = 25 + total / 2
            # else
            #    x = total

            # x = round_down(x)

            # if x > 99
            # overall_player = 99
            # else
            # overall_player = x

            # GOALIE:

            # total = round_down(agility * 4.5) + round_down(defensive * 4.5) + round_down(puck_ctl * 4.5)
            # + (stick_r * 1) + (stick_l * 1) + (glove_r * 1) + (glove_l * 1)

            # if total < 50
            # x = 25 + total / 2
            # else
            # x = total

            # x = rounddown(x)

            # if x > 99
            # overall_goalie = 99
            # else
            # overall_goalie = x

            if pos == 'G':
                total = int(attrib[1] * 4.5) + int(attrib[4] * 4.5) + int(attrib[5] * 4.5) + attrib[10] + attrib[11] \
                        + attrib[12] + attrib[13]
                if total < 50:
                    ovr = int(25 + (total / 2))
                else:
                    ovr = total

            else:
                total = (attrib[1] * 2) + (attrib[2] * 3) + (attrib[3] * 3) + (attrib[4] * 2) + attrib[5] + \
                        (attrib[6] * 2) + (attrib[8] * 3) + (attrib[9] * 2) + attrib[10] + attrib[12]
                if total < 50:
                    ovr = int(25 + (total / 2))
                else:
                    ovr = total

            if ovr > 99:
                ovr = 99

            # Print

            output = {"First": names[0], "Last": names[1], "Abv": tminfo['abv'], "Pos": pos, "JNo": jno, "Ovr": ovr,
                      "Wgt": attrib[0],
                      "Agl": attrib[1], "Spd": attrib[2], "OfA": attrib[3], "DfA": attrib[4], "ShP-PkC": attrib[5],
                      "Chk": attrib[6], "Hnd": attrib[7], "StH": attrib[8], "ShA": attrib[9], "End-StR": attrib[10],
                      "Rgh-StL": attrib[11], "Pas-GlR": attrib[12], "Agr-GlL": attrib[13]}
            writer.writerow(output)


def main():
    root = Tk()
    ros = RosExt(root)

    # Window Setting
    root.geometry("300x250+300+300")
    root.resizable(False, False)
    root.wm_iconbitmap('icon.ico')
    root.mainloop()


if __name__ == '__main__':
    main()
