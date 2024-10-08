"""
Extracts nested contacts data from Cellebrite formatted Excel documents.
    - Cellebrite Stores contact details in multiline Excel cells.

Formatted unapologetically with Black

# Current Known Issues
# FIXME - Fix Instagram parser, account ID's have changed
# FIXME - Fix Og Signal parser, Column order

Changelog
0.9 - Fix - Handles name change to Signal Private Messenger and extra data columns
    - prints version to command line
    - Fix bug where files ending in .XLXS (Caps) wouldn't be automatically found
    - Support for Outlook contacts

0.8 - Fix issue where Input file was entered twice for Instagram export sheet

0.7 - Add Provenance data column
    - Fix issue where WhatsApp or Facebook may not export due to lack if 'Interaction Statuses' Column

0.6 - Fix issue with Threema user ID attribution
    - Fix issue with parsers crashing out, now raises an exception and continues.

0.5 - Added support for recents - at this time this is kept separate from native contacts
    - Warning re large files, pandas is unable to provide load time estimates
    - Add option to normalise Au mobile phone by converting +614** to 04**
    - Minor tidyups and fixes to logging.
    - Fix WeChat exception for older style excels
    - Fix Whatsapp exception when interaction status is not populated
    - Fix Exception when there is no IMEI entry at all, eg. older iPads
    - Populate and export source columns

0.4a - Added support for Cellebrite files with device info stored in "device" rather than name columns

0.4 - Add support for alternate Cellebrite info page format
    - Add support For Line, WeChat, Threema contacts

0.3 Complete rewrite

0.2 - Implement command line argument parser
        Allow bulk processing of all items in directory

0.1 - Initial concept

"""

import argparse
import glob
import logging
import os
import pandas as pd
from pathlib import Path
import sys

## Details
__description__ = 'Flattens Cellebrite formatted Excel files. "Contacts" and "Device Info" tabs are required.'
__author__ = "facelessg00n"
__version__ = "0.9"

parser = argparse.ArgumentParser(
    description=__description__,
    epilog="Developed by {}, version {}".format(str(__author__), str(__version__)),
)

# ----------- Options -----------
os.chdir(os.getcwd())

# Show extra debug output
debug = False

# Normalise Australian mobile numbers by replacing +614 with 04
ausNormal = True

# File size warning (MB)
warnSize = 50


# ----------- Logging options -------------------------------------

logging.basicConfig(
    filename="clbExtract.log",
    format="%(asctime)s,- %(levelname)s - %(message)s",
    level=logging.INFO,
)


# Set names for sheets of interest
clbPhoneInfo = "Device Info"
clbContactSheet = "Contacts"
clbPhoneInfov2 = "Device Information"

# FIXME
#### ---- Column names and other options ---------------------------------------------
provenanceCols = ["WARRANT", "COLLECT", "EXAM", "NOTICE"]

contactOutput = "ContactDetail"
contactTypeOutput = "ContactType"
originIMEI = "originIMEI"
parsedApps = [
    "Facebook Messenger",
    "Instagram",
    "Line",
    "Native",
    "Outlook",
    "Recents",
    "Signal",
    "Signal Private Messenger",
    "Snapchat",
    "WhatsApp",
    "Telegram",
    "Threema",
    "WeChat",
    "Zalo",
]


# Class object to hold phone and input file info
class phoneData:
    IMEI = None
    IMEI2 = None
    inFile = None
    inPath = None
    inProvenance = None

    def __init__(
        self, IMEI=None, IMEI2=None, inFile=None, inPath=None, inProvenance=None
    ) -> None:
        self.IMEI = IMEI
        self.IMEI2 = IMEI2
        self.inFile = inFile
        self.inPath = inPath
        self.inProvenance = inProvenance


# -------------Functions live here ------------------------------------------

# ----- Bulk Excel Processor--------------------------------------------------


# Finds and processes all excel files in the working directory.
def bulkProcessor(inputProvenance):
    FILE_PATH = os.getcwd()
    inputFiles = glob.glob("*.xlsx") + glob.glob("*.XLSX")
    print((str(len(inputFiles)) + " Excel files located. \n"))
    logging.info("Bulk processing {} files".format(str(len(inputFiles))))
    # If there are no files found exit the process.
    if len(inputFiles) == 0:
        print("No excel files located.")
        print("Exiting.")
        quit()
    else:
        for inputFile in inputFiles:
            if os.path.exists(inputFile):
                try:
                    processMetadata(inputFile, inputProvenance)
                # Need to deal with $ files.
                except FileNotFoundError:
                    print("File does not exist or temp file detected")
                    pass
    if debug:
        for inputFile in inputFiles:
            inputFilename = inputFile.split(".")[0]
            print(inputFilename)


# FIXME - Deal with error when this info is missing
### -------- Process phone metadata ------------------------------------------------------
def processMetadata(inputFile, inputProvenance):
    inputFile = inputFile
    print("Input Provenance is {}".format(inputProvenance))
    print("Extracting metadata from {}".format(inputFile))
    logging.info("Extracting metadata from {}".format(inputFile))

    phoneData.inProvenance = inputProvenance

    fileSize = os.path.getsize(inputFile) / 1048576
    if fileSize > warnSize:
        print(
            "Large input file detected, {} MB and may take some time to process, sadly progress is not able to be updated while the file is loading".format(
                f"{fileSize:.2f}"
            )
        )
    else:
        print("Input file is {} MB".format(f"{fileSize:.2f}"))

    try:
        infoPD = pd.read_excel(
            inputFile, sheet_name=clbPhoneInfo, header=1, usecols="B,C,D"
        )

        try:
            phoneData.IMEI = infoPD.loc[infoPD["Name"] == "IMEI", ["Value"]].values[0][
                0
            ]
            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)
            phoneData.inProvenance = inputProvenance
        except:
            print("Attempting Device Column")
            try:
                phoneData.IMEI = infoPD.loc[
                    infoPD["Device"] == "IMEI", ["Value"]
                ].values[0][0]
                phoneData.inFile = Path(inputFile).stem
                phoneData.inPath = os.path.dirname(inputFile)
            except:
                print("IMEI not located, setting to NULL")
                phoneData.IMEI = None
                phoneData.inFile = Path(inputFile).stem
                phoneData.inPath = os.path.dirname(inputFile)

        try:
            phoneData.IMEI2 = infoPD.loc[infoPD["Name"] == "IMEI2", ["Value"]].values[
                0
            ][0]
        except:
            phoneData.IMEI2 = None
            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)
        # phoneData.inFile = inputFile.split(".")[0]
        phoneData.inFile = Path(inputFile).stem
        phoneData.inPath = os.path.dirname(inputFile)

        if debug:
            print(infoPD)
            print(phoneData.IMEI)

    except ValueError:
        print(
            "Info tab not found in {}, attempting with second format.".format(inputFile)
        )
        logging.exception(
            "No info tab found in {}, attempting with second format".format(inputFile)
        )
        try:
            infoPD = pd.read_excel(
                inputFile, sheet_name=clbPhoneInfov2, header=1, usecols="B,C,D"
            )
            # Remove leading whitespace from columns
            infoPD["Name"] = infoPD["Name"].str.strip()
            phoneData.IMEI = infoPD.loc[infoPD["Name"] == "IMEI", ["Value"]].values[0][
                0
            ]
            print("Second format succeeded")
            logging.info("Second format succeeded on {}".format(inputFile))

            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)

        except IndexError:
            print("IMEI not located, is this a tablet or iPAD?")
            logging.warning(
                "IMEI not found in {}, apptempting with with no IMEI".format(inputFile)
            )
            phoneData.IMEI = None
            phoneData.IMEI2 = None
            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)
            print("Loaded {}, with no IMEI".format(inputFile))
            logging.info("Loaded {}, with no IMEI".format(inputFile))
            pass

        except ValueError:
            print(
                "\033[1;31m Info tab not found in {}, apptempting with with no IMEI".format(
                    inputFile
                )
            )
            logging.warning(
                "Info tab not found in {}, apptempting with with no IMEI".format(
                    inputFile
                )
            )
            phoneData.IMEI = None
            phoneData.IMEI2 = None
            # phoneData.inFile = inputFile.split(".")[0]
            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)
            print("\033[1;31m Loaded {}, with no IMEI".format(inputFile))
            logging.info("Loaded {}, with no IMEI".format(inputFile))
            pass

    try:
        processContacts(inputFile)
    except Exception as e:
        print(e)
    except ValueError:
        print("\033[1;31m No Contacts tab  found, is this a correctly formatted Excel?")
        logging.error(
            "No Contacts tab  found in {}, is this a correctly formatted Excel?".format(
                inputFile
            )
        )


### Extract contacts tab of Excel file -------------------------------------------------------------------
# This creates the initial dataframe, future processing is from copies of this dataframe.
def processContacts(inputFile):
    inputFile = inputFile
    fileSize = os.path.getsize(inputFile) / 1048576
    print("Processing contacts in {} has begun.".format(phoneData.inFile))
    logging.info("Processing contacts in {} has begun.".format(phoneData.inFile))

    if fileSize > warnSize:
        print(
            "Large input file detected, {} MB and may take some time to process, sadly progress is not able to be updated while the file is loading".format(
                f"{fileSize:.2f}"
            )
        )
    else:
        print("Input file is {} MB".format(f"{fileSize:.2f}"))

    # Record input filename for use in export processes.
    if debug:
        print("\033[0;37m Input file is : {}".format(phoneData.inFile))

    contactsPD = pd.read_excel(
        inputFile,
        sheet_name=clbContactSheet,
        header=1,
        index_col="#",
        usecols=["#", "Name", "Entries", "Source", "Account"],
    )

    print("\033[0mProcessing the following app types for : {}".format(phoneData.inFile))
    applist = contactsPD["Source"].unique()
    for x in applist:
        if x in parsedApps:
            print("{} : \u2713 ".format(x))

        else:
            print("{} : \u2716".format(x))

    # Process native contacts
    try:
        processAppleNative(contactsPD)
    except Exception as e:
        print("Processing native contacts failed")
        print(e)
        pass

    # Process Apps
    for x in applist:
        if x == "Facebook Messenger":
            try:
                processFacebookMessenger(contactsPD)
            except Exception as e:
                logging.warning("Failed to parse Facebook messenger - {}".format(e))
                pass
        if x == "Instagram":
            try:
                processInstagram(contactsPD)
            except:
                logging.warning("Failed to parse Instagram")
                pass
        if x == "Line":
            try:
                processLine(contactsPD)
            except:
                logging.warning("Failed to parse Line")
                pass
        if x == "Outlook":
            try:
                processOutlookContacts(contactsPD)
            except:
                logging.warning("Failed to parse Outlook")
                pass
        if x == "Recents":
            try:
                processRecents(contactsPD)
            except:
                logging.warning("Failed to parse Recents")
                pass
        if x == "Snapchat":
            try:
                processSnapChat(contactsPD)
            except:
                logging.warning("Failed to parse Snapchat")
                pass
        if x == "Telegram":
            try:
                processTelegram(contactsPD)
            except:
                logging.warning("Failed to parse Telegram")
                pass
        if x == "Threema":
            try:
                processThreema(contactsPD)
            except:
                logging.warning("Failed to parse Threema")
                pass
        if x == "Signal":
            try:
                processSignal(contactsPD)
            except:
                logging.warning("Failed to parse Signal")
                pass
        if x == "Signal Private Messenger":
            try:
                processSignalPrivateMessenger(contactsPD)
            except:
                logging.warning("Failed to parse Signal Private Messenger")

        if x == "WeChat":
            try:
                processWeChat(contactsPD)
            except:
                logging.warning("Failed to parse WeChat")
                pass
        if x == "WhatsApp":
            try:
                processWhatsapp(contactsPD)
            except Exception as e:
                logging.warning("Failed to parse WhatsApp - {}".format(e))
                pass
        if x == "Zalo":
            try:
                processZalo(contactsPD)
            except:
                logging.warning("Failed to parse Zalo")
                pass

    print("\nProcessing of {} complete".format(inputFile))


# ------ Parse Facebook Messenger --------------------------------------------------------------
def processFacebookMessenger(contactsPD):
    print("\nProcessing Facebook Messenger")
    facebookMessengerPD = contactsPD[contactsPD["Source"] == "Facebook Messenger"]
    facebookMessengerPD = facebookMessengerPD.drop("Entries", axis=1).join(
        facebookMessengerPD["Entries"].str.split("\n", expand=True)
    )
    facebookMessengerPD = facebookMessengerPD.reset_index(drop=True)

    selected_cols = []
    for x in facebookMessengerPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def phoneCheck(facebookMessengerPD):
        for x in selected_cols:
            facebookMessengerPD.loc[
                (facebookMessengerPD[x].str.contains("User ID-Facebook Id", na=False)),
                "Account ID",
            ] = facebookMessengerPD[x].str.split(":", n=1, expand=True)[1]
            facebookMessengerPD.loc[
                (facebookMessengerPD[x].str.contains("User ID-Username", na=False)),
                "User Name",
            ] = facebookMessengerPD[x].str.split(":", n=1, expand=True)[1]

    phoneCheck(facebookMessengerPD)

    facebookMessengerPD["Source"] = "Messenger"
    facebookMessengerPD[originIMEI] = phoneData.IMEI
    facebookMessengerPD["inputFile"] = phoneData.inFile
    facebookMessengerPD["Provenance"] = phoneData.inProvenance

    exportCols = []
    for x in facebookMessengerPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    print(
        "{} user accounts located".format(len(facebookMessengerPD["Account"].unique()))
    )
    print("{} contacts located".format(len(facebookMessengerPD["Account ID"].unique())))
    print("Exporting {}-FB-MESSENGER.csv".format(phoneData.inFile))
    logging.info("Exporting FB messenger from {}".format(phoneData.inFile))
    try:
        facebookMessengerPD[exportCols].to_csv(
            "{}-FB-MESSENGER.csv".format(phoneData.inFile),
            index=False,
        )
    except Exception as e:
        print(e)


# ----- Parse Instagram data ------------------------------------------------------------------
def processInstagram(contactsPD):
    print("\nProcessing Instagram")
    instagramPD = contactsPD[contactsPD["Source"] == "Instagram"].copy()
    instagramPD = instagramPD.drop("Entries", axis=1).join(
        instagramPD["Entries"].str.split("\n", expand=True)
    )

    selected_cols = []
    for x in instagramPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def instaContacts(instagramPD):
        for x in selected_cols:
            instagramPD.loc[
                (instagramPD[x].str.contains("User ID-Username", na=False)), "User Name"
            ] = instagramPD[x].str.split(":", n=1, expand=True)[1]
            instagramPD.loc[
                (instagramPD[x].str.contains("User ID-Instagram Id", na=False)),
                "Instagram ID",
            ] = instagramPD[x].str.split(":", n=1, expand=True)[1]

    instaContacts(instagramPD)

    instagramPD[originIMEI] = phoneData.IMEI
    instagramPD["inputFile"] = phoneData.inFile

    exportCols = []
    for x in instagramPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    print("{} Instagram contacts located".format(len(instagramPD["Name"])))
    print("Exporting {}-INSTAGRAM.csv".format(phoneData.inFile))
    logging.info("Exporting Instagram from {}".format(phoneData.inFile))
    # TODO - Fix column handling
    instagramPD[exportCols].to_csv(
        "{}-INSTAGRAM.csv".format(phoneData.inFile),
        index=False,
    )


# ---- Process Line -----------------------------------------------------------------------
def processLine(contactsPD):
    print("Processing Line")
    linePD = contactsPD[contactsPD["Source"] == "Line"].copy()
    linePD = linePD.drop("Entries", axis=1).join(
        linePD["Entries"].str.split("\n", expand=True)
    )
    linePD = linePD.reset_index(drop=True)

    selected_cols = []
    for x in linePD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def processLine(LinePD):
        for x in selected_cols:
            LinePD.loc[
                (LinePD[x].str.contains("User ID-Address Book Name:", na=False)),
                "LineAddressBook",
            ] = LinePD[x].str.split(":", n=1, expand=True)[1]

            LinePD.loc[
                (LinePD[x].str.contains("User ID-User ID:", na=False)),
                "LineUserID",
            ] = LinePD[x].str.split(":", n=1, expand=True)[1]
            LinePD.loc[
                (LinePD[x].str.contains("User ID-Server:", na=False)),
                "LineServerID",
            ] = LinePD[x].str.split(":", n=1, expand=True)[1]

    processLine(linePD)

    linePD[originIMEI] = phoneData.IMEI
    linePD["inputFile"] = phoneData.inFile
    exportCols = []

    for x in linePD.columns:
        if isinstance(x, str):
            exportCols.append(x)

    print("{} Line contacts located".format(len(linePD["Name"])))
    print("Exporting {}-LINE.csv".format(phoneData.inFile))
    logging.info("Exporting Line contacts from {}".format(phoneData.inFile))
    linePD[exportCols].to_csv("{}-LINE.csv".format(phoneData.inFile), index=False)


# ------------Process native contact list ------------------------------------------------
def processAppleNative(contactsPD):

    print("\nProcessing Native Contacts")
    # nativeContactsPD = contactsPD[contactsPD["Source"].isna()]

    # Contacts are stored with either null (iPhone) or "Phone" for Android
    nativeContactsPD = contactsPD[
        (contactsPD.Source.isna()) | (contactsPD.Source == "Phone")
    ].copy()

    # Fill NaN values with : to prevent error with blank entries.
    nativeContactsPD.Entries = nativeContactsPD.Entries.fillna(":")

    nativeContactsPD = nativeContactsPD.drop("Entries", axis=1).join(
        nativeContactsPD["Entries"]
        .str.split("\n", expand=True)
        .stack()
        .reset_index(level=1, drop=True)
        .rename("Entries")
    )

    # nativeContactsPD = nativeContactsPD[["Name", "Interaction Statuses", "Entries"]]

    nativeContactsPD = nativeContactsPD[
        nativeContactsPD["Entries"].str.contains(r"Phone-")
    ]
    nativeContactsPD[originIMEI] = phoneData.IMEI
    nativeContactsPD["inputFile"] = phoneData.inFile
    nativeContactsPD["Provenance"] = phoneData.inProvenance

    # Remove erroneous characters, need to make this a regex
    # TODO Use a regex to tidy this up.
    nativeContactsPD["Entries"] = (
        nativeContactsPD["Entries"]
        .str.split(":", n=1, expand=True)[1]
        .str.strip()
        .str.replace(" ", "", regex=False)
        .str.replace("-", "", regex=False)
        .str.replace("+", "", regex=False)
        # Fix issue with inseyets reports
        .str.replace("Message", "", regex=False)
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
    )

    if ausNormal:
        nativeContactsPD["Entries"] = nativeContactsPD["Entries"].str.replace(
            r"\+614", "04", regex=True
        )

    if debug:
        print(nativeContactsPD)

    # nativeContactsPD = nativeContactsPD[[originIMEI, "Name", "Entries", "Interaction Statuses"]]
    print("{} contacts located.".format(len(nativeContactsPD)))
    print("Exporting {}-NATIVE.csv".format(phoneData.inFile))
    logging.info("Exporting Native contacts from {}".format(phoneData.inFile))
    nativeContactsPD.to_csv("{}-NATIVE.csv".format(phoneData.inFile), index=False)


# Process Outlook Contacts
def processOutlookContacts(contactsPD):
    print("\nProcessing Outlook Contacts")

    outlookContactsPD = contactsPD[(contactsPD.Source == "Outlook")].copy()
    # Fill NaN values with : to prevent error with blank entries.
    outlookContactsPD.Entries = outlookContactsPD.Entries.fillna(":")

    outlookContactsPD = outlookContactsPD.drop("Entries", axis=1).join(
        outlookContactsPD["Entries"]
        .str.split("\n", expand=True)
        .stack()
        .reset_index(level=1, drop=True)
        .rename("Entries")
    )

    outlookContactsPD = outlookContactsPD[["Account", "Name", "Entries", "Source"]]
    outlookContactsPD[originIMEI] = phoneData.IMEI
    outlookContactsPD["inputFile"] = phoneData.inFile
    outlookContactsPD["Provenance"] = phoneData.inProvenance

    outlookContactsPD["Entries"] = (
        outlookContactsPD["Entries"].str.split(":", n=1, expand=True)[1].str.strip()
    )

    print("{} contacts located.".format(len(outlookContactsPD)))
    print("Exporting {}-OUTLOOK.csv".format(phoneData.inFile))
    logging.info("Exporting Native contacts from {}".format(phoneData.inFile))
    outlookContactsPD.to_csv("{}-OUTLOOK.csv".format(phoneData.inFile), index=False)


# ----------- Parse Recents -----------------------------------------------------------------------
def processRecents(contactsPD):
    print("\nProcessing Recents")
    recentsPD = contactsPD[contactsPD["Source"] == "Recents"].copy()
    recentsPD.Entries = recentsPD.Entries.fillna(":")
    recentsPD = recentsPD[recentsPD["Entries"].str.contains(r"Phone-")]

    recentsPD[originIMEI] = phoneData.IMEI
    recentsPD["inputFile"] = phoneData.inFile
    recentsPD["Provenance"] = phoneData.inProvenance

    recentsPD["Entries"] = (
        recentsPD["Entries"]
        .str.split(":", n=1, expand=True)[1]
        .str.strip()
        .str.replace(" ", "")
        .str.replace("-", "")
        # .str.replace("+","",regex=False)
    )
    if ausNormal:
        recentsPD["Entries"] = recentsPD["Entries"].str.replace(
            r"\+614", "04", regex=True
        )

    print("{} recent contacts located.".format(len(recentsPD)))
    print("Exporting {}-RECENT.csv".format(phoneData.inFile))
    logging.info("Exporting recent contacts from {}".format(phoneData.inFile))
    recentsPD.to_csv("{}-RECENTS.csv".format(phoneData.inFile), index=False)


# ------------Parse Signal contacts ---------------------------------------------------------------
def processSignal(contactsPD):
    print("\nProcessing Signal Contacts")
    signalPD = contactsPD[contactsPD["Source"] == "Signal"].copy()
    signalPD = signalPD[["Name", "Entries", "Source"]]
    signalPD = signalPD.drop("Entries", axis=1).join(
        signalPD["Entries"].str.split("\n", expand=True)
    )

    # Data is expended into columns with integer names, add these columsn to selected_cols so we can search them later
    selected_cols = []
    for x in signalPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    # FIXME improve with method used for other apps
    # Signal can store mutiple values under entries such as Mobile Number:
    # So we break them all out into columns.
    def signalContact(signalPD):
        for x in selected_cols:
            # Locate Signal Username and move to Username Column
            signalPD.loc[
                (signalPD[x].str.contains("User ID-Username:", na=False)),
                "User Name",
            ] = signalPD[x].str.split(":", n=1, expand=True)[1]
            # Delete Username entry from origional location
            signalPD.loc[
                signalPD[x].str.contains("User ID-Username:", na=False), [x]
            ] = ""
            # delete all befote semicolon
            signalPD[x] = signalPD[x].str.split(":", n=1, expand=True)[1].str.strip()

    signalContact(signalPD)

    signalPD[originIMEI] = phoneData.IMEI
    signalPD["inputFile"] = phoneData.inFile
    signalPD["Provenance"] = phoneData.inProvenance

    export_cols = [originIMEI, "Name", "User Name"]
    export_cols.extend(selected_cols)
    print("Located {} Signal contacts".format(len(signalPD["Name"])))
    print("Exporting {}-SIGNAL.csv".format(phoneData.inFile))
    logging.info("Exporting Signal messenger from {}".format(phoneData.inFile))
    signalPD.to_csv(
        "{}-SIGNAL.csv".format(phoneData.inFile), index=False, columns=export_cols
    )


# ----------- Parse Signal Private Messenger--------------------------------------------------------
def processSignalPrivateMessenger(contactsPD):
    print("\nProcessing Signal Private Messenger")
    spmPD = contactsPD[contactsPD["Source"] == "Signal Private Messenger"].copy()
    spmPD = spmPD.drop("Entries", axis=1).join(
        spmPD["Entries"].str.split("\n", expand=True)
    )
    # spmPD['Entries'].tolist()
    # spmPD.explode('Entries')
    # spmPD = spmPD.reset_index(drop=True)
    spmPD[originIMEI] = phoneData.IMEI
    spmPD["inputFile"] = phoneData.inFile
    spmPD["Provenance"] = phoneData.inProvenance

    selected_cols = []
    for x in spmPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def spmContacts(spmPD):
        for x in selected_cols:
            try:
                spmPD.loc[(spmPD[x].str.contains("Phone-:", na=False)), "Phone"] = (
                    spmPD[x].str.split(":", n=1, expand=True)[1]
                )
            except:
                pass
            try:
                spmPD.loc[(spmPD[x].str.contains("User ID-:", na=False)), "User-ID"] = (
                    spmPD[x].str.split(":", n=1, expand=True)[1]
                )
            except:
                pass
            try:
                spmPD.loc[
                    (spmPD[x].str.contains("User ID-Nickname:", na=False)),
                    "User-ID-Nickname",
                ] = spmPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass
            try:
                spmPD.loc[
                    (spmPD[x].str.contains("User ID-Username:", na=False)),
                    "User-ID-Username",
                ] = spmPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass
            try:
                spmPD.loc[
                    (spmPD[x].str.contains("User ID-ProfileKey:", na=False)),
                    "User-ID-ProfileKey",
                ] = spmPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

    spmContacts(spmPD)
    # spmPD.info()

    exportCols = []
    # Remove column from previous step
    for x in spmPD.columns:
        if isinstance(x, str):
            # if x !="Provenence" or x != 'originIMEI' or x != 'inputFile':
            if x not in ["Provenance", "originIMEI", "inputFile"]:
                exportCols.append(x)

    exportCols.extend(["originIMEI", "inputFile", "Provenance"])

    print("Located {} Signal Private Messenger contacts.".format(len(spmPD["Name"])))
    print("Exporting {}-Signal-PM.csv".format(phoneData.inFile))
    logging.info("Exporting Signal Private Messenger from {}".format(phoneData.inFile))
    spmPD[exportCols].to_csv("{}-Signal-PM.csv".format(phoneData.inFile), index=False)


# ----------- Parse Snapchat data ------------------------------------------------------------------
def processSnapChat(contactsPD):
    print("\nProcessing Snapchat")
    snapPD = contactsPD[contactsPD["Source"] == "Snapchat"]
    snapPD = snapPD[["Name", "Entries", "Source"]]

    # Extract nested entities
    snapPD = snapPD.drop("Entries", axis=1).join(
        snapPD["Entries"].str.split("\n", expand=True)
    )
    selected_cols = []
    for x in snapPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def snapContacts(snapPD):
        for x in selected_cols:
            snapPD.loc[
                (snapPD[x].str.contains("User ID-Username", na=False)), "User Name"
            ] = snapPD[x].str.split(":", n=1, expand=True)[1]
            snapPD.loc[
                (snapPD[x].str.contains("User ID-User ID", na=False)), "User ID"
            ] = snapPD[x].str.split(":", n=1, expand=True)[1]

    snapContacts(snapPD)

    snapPD[originIMEI] = phoneData.IMEI
    snapPD["inputFile"] = phoneData.inFile
    snapPD["Provenance"] = phoneData.inProvenance

    exportCols = []
    for x in snapPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    if debug:
        print(snapPD[exportCols])

    print("{} Snapchat contacts located.".format(len(snapPD)))
    print("Exporting {}-SNAPCHAT.csv".format(phoneData.inFile))
    logging.info("Exporting Snapchat from {}".format(phoneData.inFile))
    snapPD[exportCols].to_csv(
        "{}-SNAPCHAT.csv".format(phoneData.inFile),
        index=False,
        columns=[
            originIMEI,
            "Name",
            "User Name",
            "User ID",
            "Source",
            "inputFile",
            "Provenance",
        ],
    )


# ---- Parse Telegram Contacts--------------------------------------------------------------
def processTelegram(contactsPD):
    print("\nProcessing Telegram")
    telegramPD = contactsPD[contactsPD["Source"] == "Telegram"].copy()
    telegramPD = telegramPD.drop("Entries", axis=1).join(
        telegramPD["Entries"].str.split("\n", expand=True)
    )
    telegramPD = telegramPD.reset_index(drop=True)

    selected_cols = []
    for x in telegramPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def phoneCheck(telegramPD):
        for x in selected_cols:
            telegramPD.loc[
                (telegramPD[x].str.contains("Phone-", na=False)), "Phone-Number"
            ] = telegramPD[x].str.split(":", n=1, expand=True)[1]

            telegramPD.loc[
                (telegramPD[x].str.contains("User ID-Peer", na=False)), "Peer-ID"
            ] = telegramPD[x].str.split(":", n=1, expand=True)[1]

            telegramPD.loc[
                (telegramPD[x].str.contains("User ID-Username", na=False)), "User-Name"
            ] = telegramPD[x].str.split(":", n=1, expand=True)[1]

    phoneCheck(telegramPD)

    telegramPD[originIMEI] = phoneData.IMEI
    telegramPD["inputFile"] = phoneData.inFile
    telegramPD["Provenance"] = phoneData.inProvenance
    telegramPD["source"] = "Telegram"
    exportCols = []
    for x in telegramPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    # Export CSV
    print("{} Telegram contacts located.".format(len(telegramPD)))
    print("Exporting {}-TELEGRAM.csv".format(phoneData.inFile))
    logging.info("Exporting Telegram from {}".format(phoneData.inFile))
    telegramPD[exportCols].to_csv(
        "{}-TELEGRAM.csv".format(phoneData.inFile), index=False
    )


# ------ Parse Threema Contacts -----------------------------------------------------------------
def processThreema(contactsPD):
    print("\nProcessing Threema")
    threemaPD = contactsPD[contactsPD["Source"] == "Threema"].copy()
    threemaPD = threemaPD.drop("Entries", axis=1).join(
        threemaPD["Entries"].str.split("\n", expand=True)
    )
    threemaPD = threemaPD.reset_index(drop=True)

    selected_cols = []
    for x in threemaPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def ThreemaParse(ThreemaPD):
        for x in selected_cols:
            try:
                ThreemaPD.loc[
                    (ThreemaPD[x].str.contains("User ID-identity:", na=False)),
                    "Threema ID",
                ] = ThreemaPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass
            try:
                ThreemaPD.loc[
                    (ThreemaPD[x].str.contains("User ID-Username:", na=False)),
                    "ThreemaUsername",
                ] = ThreemaPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

    ThreemaParse(threemaPD)

    threemaPD[originIMEI] = phoneData.IMEI
    threemaPD["inputFile"] = phoneData.inFile
    threemaPD["Provenance"] = phoneData.inProvenance

    exportCols = []
    for x in threemaPD.columns:
        if isinstance(x, str):
            exportCols.append(x)

    print("Exporting {}-THREEMA.csv".format(phoneData.inFile))
    logging.info("Exporting Threema from {}".format(phoneData.inFile))
    threemaPD[exportCols].to_csv("{}-THREEMA.csv".format(phoneData.inFile), index=False)


## Parse WeChat Contacts ------------------------------------------------------------------------
def processWeChat(contactsPD):
    print("\nProcessing WeChat")
    WeChatPD = contactsPD[contactsPD["Source"] == "WeChat"].copy()
    WeChatPD = WeChatPD.drop("Entries", axis=1).join(
        WeChatPD["Entries"].str.split("\n", expand=True)
    )

    WeChatPD = WeChatPD.reset_index(drop=True)

    selected_cols = []
    for x in WeChatPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def WeChatContacts(WeChatPD):

        for x in selected_cols:
            # FIXME Usernames that contain @stranger???
            # FIXME Try / Except / Pass

            try:
                WeChatPD.loc[
                    (WeChatPD[x].str.contains("User ID-WeChat ID:", na=False)),
                    "WeChatID",
                ] = WeChatPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

            try:
                WeChatPD.loc[
                    (WeChatPD[x].str.contains("User ID-QQ:", na=False)), "QQ User ID"
                ] = WeChatPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

            try:
                WeChatPD.loc[
                    (WeChatPD[x].str.contains("User ID-Username:", na=False)),
                    "Username",
                ] = WeChatPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

            try:
                WeChatPD.loc[
                    (WeChatPD[x].str.contains("User ID-LinkedIn ID:", na=False)),
                    "LinkedIn ID",
                ] = WeChatPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

            try:
                WeChatPD.loc[
                    (WeChatPD[x].str.contains("User ID-Facebook ID:", na=False)),
                    "Facebook ID",
                ] = WeChatPD[x].str.split(":", n=1, expand=True)[1]
            except:
                pass

    WeChatContacts(WeChatPD)

    # Repalace we chat ID's with @ stranhger with blank values as are not we chat user IDs
    try:
        WeChatPD.WeChatID = WeChatPD.WeChatID.apply(
            lambda x: "" if (r"@stranger") in str(x) else x
        )
    except:
        print("WeChat float exception")
        print(WeChatPD.WeChatID)
        pass

    WeChatPD[originIMEI] = phoneData.IMEI
    WeChatPD["inputFile"] = phoneData.inFile
    WeChatPD["Provenance"] = phoneData.inProvenance
    WeChatPD["Source"] = "Weixin"

    # Export Columns where the title is a string to drop working columns
    exportCols = []
    for x in WeChatPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    print("Located {} WeChat contacts.".format(len(WeChatPD["WeChatID"])))
    print("Exporting {}-WECHAT.csv".format(phoneData.inFile))
    logging.info("Exporting WeChat from {}".format(phoneData.inFile))
    WeChatPD[exportCols].to_csv("{}-WECHAT.csv".format(phoneData.inFile), index=False)


# ---Parse Whatsapp Contacts----------------------------------------------------------------------
# Load WhatsApp
def processWhatsapp(contactsPD):
    print("\nProcessing WhatsApp")
    whatsAppPD = contactsPD[contactsPD["Source"] == "WhatsApp"].copy()
    try:
        whatsAppPD = whatsAppPD[["Name", "Entries", "Source", "Interaction Statuses"]]
        # Datatype needs to be object not float to allow filtering by string without throwing an error
        whatsAppPD["Interaction Statuses"] = whatsAppPD["Interaction Statuses"].astype(
            object
        )
        # Shared contacts are not associated with a Whats app ID and cause problems.
        print(whatsAppPD.dtypes)
        whatsAppPD = whatsAppPD[
            whatsAppPD["Interaction Statuses"].str.contains("Shared", na=False) == False
        ]
    except Exception as e:
        print(e)
        print("Interaction statuses column not found, ignoring")
        # print(whatsAppPD)
        whatsAppPD = whatsAppPD[
            [
                "Name",
                "Entries",
                "Source",
            ]
        ]

    # Unpack nested data
    whatsAppPD = whatsAppPD.drop("Entries", axis=1).join(
        whatsAppPD["Entries"].str.split("\n", expand=True)
    )

    # Data is expanded into colums with Integer names, check for these columns and add them to a
    # list to allow for different width sheets.
    colList = list(whatsAppPD)
    selected_cols = []
    for x in colList:
        if isinstance(x, int):
            selected_cols.append(x)

    # Look for data across expanded columns and shift it to output columns.
    def whatsappContactProcess(whatsAppPD):
        for x in selected_cols:
            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("Phone-Mobile", na=False)), "Phone-Mobile"
            ] = (
                whatsAppPD[x]
                .str.split(":", n=1, expand=True)[1]
                .str.replace(" ", "")
                .str.replace("-", "")
            )

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("Phone-:", na=False)), "Phone"
            ] = (
                whatsAppPD[x]
                .str.split(":", n=1, expand=True)[1]
                .str.replace(" ", "")
                .str.replace("-", "")
            )

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("Phone-Home:", na=False)), "Phone-Home"
            ] = (
                whatsAppPD[x]
                .str.split(":", n=1, expand=True)[1]
                .str.replace(" ", "")
                .str.replace("-", "")
            )

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("User ID-Push Name", na=False)), "Push-ID"
            ] = whatsAppPD[x].str.split(":", n=1, expand=True)[1]

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("User ID-Id", na=False)), "Id-ID"
            ] = whatsAppPD[x].str.split(":", n=1, expand=True)[1]

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("User ID-WhatsApp User Id", na=False)),
                "WhatsApp-ID",
            ] = whatsAppPD[x].str.split(":", n=1, expand=True)[1]

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("Web address-Professional", na=False)),
                "BusinessWebsite",
            ] = whatsAppPD[x].str.split(":", n=1, expand=True)[1]

            whatsAppPD.loc[
                (whatsAppPD[x].str.contains("Email-Professional", na=False)),
                "Business-Email",
            ] = whatsAppPD[x].str.split(":", n=1, expand=True)[1]

    whatsappContactProcess(whatsAppPD)

    # Add IMEI Column
    whatsAppPD[originIMEI] = phoneData.IMEI
    whatsAppPD["inputFile"] = phoneData.inFile
    whatsAppPD["Provenance"] = phoneData.inProvenance
    whatsAppPD["Source"] = "Whatsapp"

    # Remove working columns.
    exportCols = []
    for x in whatsAppPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    if debug:
        print(exportCols)

    # Export CSV
    print("{} WhatsApp contacts located".format(len(whatsAppPD["Name"])))
    print("Exporting {}-WHATSAPP.csv".format(phoneData.inFile))
    logging.info("Exporting Whatsapp from {}".format(phoneData.inFile))
    whatsAppPD[exportCols].to_csv(
        "{}-WHATSAPP.csv".format(phoneData.inFile), index=False
    )


# --- Parse Zalo Contacts --------------------------------------------------------------------
def processZalo(contactsPD):
    print("\nProcessinf Zalo")
    ZaloPD = contactsPD[contactsPD["Source"] == "Zalo"]
    ZaloPD = ZaloPD.drop("Entries", axis=1).join(
        ZaloPD["Entries"].str.split("\n", expand=True)
    )
    selected_cols = []
    for x in ZaloPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def processZaloContacts(ZaloPD):
        for x in selected_cols:
            ZaloPD.loc[
                (ZaloPD[x].str.contains("User ID-User Name:", na=False)),
                "ZaloUserName",
            ] = ZaloPD[x].str.split(":", n=1, expand=True)[1]

            ZaloPD.loc[
                (ZaloPD[x].str.contains("User ID-Id:", na=False)),
                "ZaloUserID",
            ] = ZaloPD[x].str.split(":", n=1, expand=True)[1]

    processZaloContacts(ZaloPD)

    ZaloPD[originIMEI] = phoneData.IMEI
    ZaloPD["inputFile"] = phoneData.inFile
    ZaloPD["Provenance"] = phoneData.inProvenance

    exportCols = []
    for x in ZaloPD.columns:
        if isinstance(x, str):
            exportCols.append(x)

    print("Exporting {}-ZALO.csv".format(phoneData.inFile))
    logging.info("Exporting Zalo from {}".format(phoneData.inFile))
    ZaloPD[exportCols].to_csv("{}-ZALO.csv".format(phoneData.inFile), index=False)


# ------- Argument parser for command line arguments -----------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}".format(str(__author__), str(__version__)),
    )

    parser.add_argument(
        "-f",
        "--f",
        dest="inputFilename",
        help="Path to Excel Spreadsheet",
        required=False,
    )

    parser.add_argument(
        "-p",
        "--p",
        dest="inputProvenance",
        choices=provenanceCols,
        required=False,
    )

    parser.add_argument(
        "-b",
        "--bulk",
        dest="bulk",
        required=False,
        action="store_true",
        help="Bulk process Excel spreadsheets in working directory.",
    )

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        parser.exit()

    if args.bulk:
        print("Bulk Process")
        bulkProcessor(args.inputProvenance)

    if args.inputFilename:
        if not os.path.exists(args.inputFilename):
            print(
                "Error: '{}' does not exist or is not a file.".format(
                    args.inputFilename
                )
            )
            sys.exit(1)
        processMetadata(args.inputFilename, args.inputProvenance)
