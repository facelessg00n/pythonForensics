"""
Extracts nested contacts data from Cellebrite formatted Excel documents.
    - Cellebrite Stores contact details in multiline Excel cells.
Formatted with Black

Changelog
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
__version__ = "0.4"

parser = argparse.ArgumentParser(
    description=__description__,
    epilog="Developed by {}".format(str(__author__), str(__version__)),
)

# ----------- Options -----------
debug = False

os.chdir(os.getcwd())

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
contactOutput = "ContactDetail"
contactTypeOutput = "ContactType"
originIMEI = "originIMEI"
parsedApps = [
    "Instagram",
    "Native",
    "Telegram",
    "Snapchat",
    "WhatsApp",
    "Facebook Messenger",
    "Signal",
    "Line",
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

    def __init__(self, IMEI=None, IMEI2=None, inFile=None, inPath=None) -> None:
        self.IMEI = IMEI
        self.IMEI2 = IMEI2
        self.inFile = inFile
        self.inPath = inPath


# -------------Functions live here ------------------------------------------

# ----- Bulk Excel Processor--------------------------------------------------


# Finds and processes all exxcel files in the working directory.
def bulkProcessor():
    FILE_PATH = os.getcwd()
    inputFiles = glob.glob("*.xlsx")
    print((str(len(inputFiles)) + " Excel files located. \n"))
    # If there are no files found exit the process.
    if len(inputFiles) == 0:
        print("No excel files located.")
        print("Exiting.")
        quit()
    else:
        for x in inputFiles:
            if os.path.exists(x):
                try:
                    processMetadata(x)
                # Need to deal with $ files.
                except FileNotFoundError:
                    print("File does not exist or temp file detected")
                    pass
    if debug:
        for x in inputFiles:
            inputFilename = x.split(".")[0]
            print(inputFilename)


# FIXME - Deal with error when this info is missing
### -------- Process phone metadata ------------------------------------------------------
def processMetadata(inputFile):
    try:
        infoPD = pd.read_excel(
            inputFile, sheet_name=clbPhoneInfo, header=1, usecols="B,C,D"
        )

        phoneData.IMEI = infoPD.loc[infoPD["Name"] == "IMEI", ["Value"]].values[0][0]
        try:
            phoneData.IMEI2 = infoPD.loc[infoPD["Name"] == "IMEI2", ["Value"]].values[
                0
            ][0]
        except:
            phoneData.IMEI2 = None
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
            # print(infoPD)
            # print(infoPD.loc[infoPD["Name"] == "IMEI", ["Value"]].values[0][0])
            phoneData.inFile = Path(inputFile).stem
            phoneData.inPath = os.path.dirname(inputFile)

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

    try:
        processContacts(inputFile)
    except ValueError:
        print("\033[1;31m No Contacts tab  found, is this a correctly formatted Excel?")
        logging.error(
            "No Contacts tab  found in {}, is this a correctly formatted Excel?".format(
                inputFile
            )
        )


### Extract contacts tab of Excel file -------------------------------------------------------------------
def processContacts(inputFile):
    inputFile = inputFile
    logging.info("Processing contacts in {} has begun.".format(inputFile))

    # Record input filename for use in export processes.

    if debug:
        print("\033[0;37m Input file is : {}".format(phoneData.inFile))

    contactsPD = pd.read_excel(
        inputFile,
        sheet_name=clbContactSheet,
        header=1,
        index_col="#",
        usecols=["#", "Name", "Interaction Statuses", "Entries", "Source", "Account"],
    )

    print(
        "\033[0m Processing the following app types for : {}".format(phoneData.inFile)
    )
    applist = contactsPD["Source"].unique()
    for x in applist:
        if x in parsedApps:
            print("{} : \u2713 ".format(x))
        else:
            print("{} : \u2716".format(x))
    # Process native contacts
    try:
        processAppleNative(contactsPD)
    except:
        print("Processing native contacts failed.")
        pass
    # Process Apps
    for x in applist:
        if x == "Facebook Messenger":
            processFacebookMessenger(contactsPD)
        if x == "Instagram":
            processInstagram(contactsPD)
        if x == "Line":
            processLine(contactsPD)
        if x == "Snapchat":
            processSnapChat(contactsPD)
        if x == "Telegram":
            processTelegram(contactsPD)
        if x == "Threema":
            processThreema(contactsPD)
        if x == "Signal":
            processSignal(contactsPD)
        if x == "WeChat":
            processWeChat(contactsPD)
        if x == "WhatsApp":
            processWhatsapp(contactsPD)
        if x == "Zalo":
            processZalo(contactsPD)


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
    facebookMessengerPD[originIMEI] = phoneData.IMEI
    exportCols = []
    for x in facebookMessengerPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    print("\n")
    print(
        "{} user accounts located".format(len(facebookMessengerPD["Account"].unique()))
    )
    print("{} contacts located".format(len(facebookMessengerPD["Account ID"].unique())))
    print("Exporting {}-FB-MESSENGER.csv".format(phoneData.inFile))
    logging.info("Exporting FB messenger from {}".format(phoneData.inFile))
    facebookMessengerPD[exportCols].to_csv(
        "{}-FB-MESSENGER.csv".format(phoneData.inFile),
        index=False,
        columns=[
            originIMEI,
            "Account",
            "Interaction Statuses",
            "Name",
            "User Name",
            "Account ID",
            "Source",
        ],
    )


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
    exportCols = []
    for x in instagramPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    print("{} Instagram contacts located".format(len(instagramPD["Name"])))
    print("Exporting {}-INSTAGRAM.csv".format(phoneData.inFile))
    logging.info("Exporting Instagram from {}".format(phoneData.inFile))
    instagramPD[exportCols].to_csv(
        "{}-INSTAGRAM.csv".format(phoneData.inFile),
        index=False,
        columns=[
            originIMEI,
            "Account",
            "Name",
            "User Name",
            "Instagram ID",
            "Interaction Statuses",
        ],
    )


# ---- Process Line -----------------------------------------------------------------------
def processLine(contactsPD):
    print("Processing Line")
    LinePD = contactsPD[contactsPD["Source"] == "Line"].copy()
    LinePD = LinePD.drop("Entries", axis=1).join(
        LinePD["Entries"].str.split("\n", expand=True)
    )
    LinePD = LinePD.reset_index(drop=True)

    selected_cols = []
    for x in LinePD.columns:
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

    processLine(LinePD)
    LinePD[originIMEI] = phoneData.IMEI
    exportCols = []

    for x in LinePD.columns:
        if isinstance(x, str):
            exportCols.append(x)

    print("{} Line contacts located".format(len(LinePD["Name"])))
    print("Exporting {}-LINE.csv".format(phoneData.inFile))
    logging.info("Exporting Line contacts from {}".format(phoneData.inFile))
    LinePD[exportCols].to_csv("{}-LINE.csv".format(phoneData.inFile), index=False)


# ------------Process native contact list ------------------------------------------------
def processAppleNative(contactsPD):
    print("\nProcessing Native Contacts")
    # nativeContactsPD = contactsPD[contactsPD["Source"].isna()]

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

    nativeContactsPD = nativeContactsPD[["Name", "Interaction Statuses", "Entries"]]

    nativeContactsPD = nativeContactsPD[
        nativeContactsPD["Entries"].str.contains(r"Phone-")
    ]
    nativeContactsPD[originIMEI] = phoneData.IMEI
    nativeContactsPD["Entries"] = (
        nativeContactsPD["Entries"]
        .str.split(":", n=1, expand=True)[1]
        .str.strip()
        .str.replace(" ", "")
        .str.replace("-", "")
    )
    if debug:
        print(nativeContactsPD)
    nativeContactsPD = nativeContactsPD[
        [originIMEI, "Name", "Entries", "Interaction Statuses"]
    ]
    print("{} contacts located.".format(len(nativeContactsPD)))
    print("Exporting {}-NATIVE.csv".format(phoneData.inFile))
    logging.info("Exporting Native contacts from {}".format(phoneData.inFile))
    nativeContactsPD.to_csv("{}-NATIVE.csv".format(phoneData.inFile), index=False)


# ------------Parse Signal contacts ---------------------------------------------------------------
def processSignal(contactsPD):
    print("\nProcessing Signal Contacts")
    signalPD = contactsPD[contactsPD["Source"] == "Signal"].copy()
    signalPD = signalPD[["Name", "Entries", "Source"]]
    signalPD = signalPD.drop("Entries", axis=1).join(
        signalPD["Entries"].str.split("\n", expand=True)
    )

    # Data is expended into columns with integern names, add these columsn to selected_cols so we can search them later
    selected_cols = []
    for x in signalPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

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

    export_cols = [originIMEI, "Name", "User Name"]
    export_cols.extend(selected_cols)
    print("Located {} Signal contacts".format(len(signalPD["Name"])))
    print("Exporting {}-SIGNAL.csv".format(phoneData.inFile))
    logging.info("Exporting Signal messenger from {}".format(phoneData.inFile))
    signalPD.to_csv(
        "{}-SIGNAL.csv".format(phoneData.inFile), index=False, columns=export_cols
    )


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

    exportCols = []
    for x in snapPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    if debug:
        print(snapPD[exportCols])
    print("Exporting {}-SNAPCHAT.csv".format(phoneData.inFile))
    logging.info("Exporting Snapchat from {}".format(phoneData.inFile))
    snapPD[exportCols].to_csv(
        "{}-SNAPCHAT.csv".format(phoneData.inFile),
        index=False,
        columns=[originIMEI, "Name", "User Name", "User ID"],
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
    exportCols = []
    for x in telegramPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    # Export CSV
    print("Exporting {}-TELEGRAM.csv".format(phoneData.inFile))
    logging.info("Exporting Telegram from {}".format(phoneData.inFile))
    telegramPD[exportCols].to_csv(
        "{}-TELEGRAM.csv".format(phoneData.inFile), index=False
    )


# ------ Parse Threema Contacts -----------------------------------------------------------------
def processThreema(contactsPD):
    ThreemaPD = contactsPD[contactsPD["Source"] == "Threema"].copy()
    ThreemaPD = ThreemaPD.drop("Entries", axis=1).join(
        ThreemaPD["Entries"].str.split("\n", expand=True)
    )
    ThreemaPD = ThreemaPD.reset_index(drop=True)

    selected_cols = []
    for x in ThreemaPD.columns:
        if isinstance(x, int):
            selected_cols.append(x)

    def ThreemaParse(ThreemaPD):
        for x in selected_cols:
            ThreemaPD.loc[
                (ThreemaPD[x].str.contains("User ID-:", na=False)), "Threema ID"
            ] = ThreemaPD[x].str.split(":", n=1, expand=True)[1]

    ThreemaParse(ThreemaPD)
    ThreemaPD[originIMEI] = phoneData.IMEI

    exportCols = []
    for x in ThreemaPD.columns:
        if isinstance(x, str):
            exportCols.append(x)

    print("Exporting {}-THREEMA.csv".format(phoneData.inFile))
    logging.info("Exporting Threema from {}".format(phoneData.inFile))
    ThreemaPD[exportCols].to_csv("{}-THREEMA.csv".format(phoneData.inFile), index=False)


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
    WeChatPD.WeChatID = WeChatPD.WeChatID.apply(
        lambda x: "" if (r"@stranger") in x else x
    )

    WeChatPD[originIMEI] = phoneData.IMEI

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
    whatsAppPD = whatsAppPD[["Name", "Entries", "Source", "Interaction Statuses"]]
    # Shared contacts are not associated with a Whats app ID and cause problems.
    whatsAppPD = whatsAppPD[
        whatsAppPD["Interaction Statuses"].str.contains("Shared", na=False) == False
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
        print("\nProcessing WhatsApp")
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

    # Remove working columns.
    exportCols = []
    for x in whatsAppPD.columns:
        if isinstance(x, str):
            exportCols.append(x)
    if debug:
        print(exportCols)

    # Export CSV
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
        bulkProcessor()

    if args.inputFilename:
        if not os.path.exists(args.inputFilename):
            print(
                "Error: '{}' does not exist or is not a file.".format(
                    args.inputFilename
                )
            )
            sys.exit(1)
        processMetadata(args.inputFilename)
