"""
Extracts nested contacts data from Cellebrite formatted Excel documents.
    - Cellebrite Stores contact details in multiline Excel cells.
Formatted with Black

"""
import argparse
import glob
import logging
import os
import pandas as pd
import sys


# Details
__description__ = "Extract contact data from Cellebrite Excel files"
__author__ = "facelessg00n"
__version__ = "0.1"

# ---------------------Options----------------------------------------------
debug = False
os.chdir(os.getcwd())

clbInfoSheet = "Summary"
clbContactSheet = "Contacts"

# Names for output columns.
contactOutput = "contactDetail1"
contactTypeOutput = "contactType"

# -----------------Exception classes----------------------------------------
class Error(Exception):
    """Base path for other exceptions"""

    pass


class FormatError(Error):
    print("Format error")
    pass


# ---------------------Data--------------------------------------------------
List_of_Columns = [
    "extractionGUID",
    "caseNumber",
    "Name",
    contactTypeOutput,
    contactOutput,
    "Group",
    "Interaction Statuses",
    "Notes",
    "Organizations",
    "Addresses",
    "Times contacted",
    "Account",
    "Deleted",
    "Tag Note",
    "Entries",
    "InputFile",
]

# ----------------------Config the logger------------------------------------
logging.basicConfig(
    filename="log.txt",
    format="%(levelname)s:%(asctime)s:%(message)s",
    level=logging.DEBUG,
)

# ---------------------For future use----------------------------------------
# FIXME finish command line input.
parser = argparse.ArgumentParser(
    description=__description__,
    epilog="Developed by {}".format(str(__author__), str(__version__)),
)

# ---------------------Input Files ------------------------------------------
# Test file
# inputFile = "phone.xlsx"

FILE_PATH = os.getcwd()
inputFiles = glob.glob("*.xlsx")

print((str(len(inputFiles)) + " Excel files located. \n"))
# If there are no files foind exit the process.
if len(inputFiles) == 0:
    print("No excel files located.")
    print("Exiting.")
    quit()
if debug:
    for x in inputFiles:
        inputFilename = x.split(".")[0]
        print(inputFilename)

# ---------------Create output folder----------------------------------------

dirName = "OUTPUT_FILES"
try:
    os.mkdir(dirName)
    print("Folder " + str(dirName) + " created")
except FileExistsError:
    print("Output directory already exists")
OUT_DIRECTORY = os.path.abspath("OUTPUT_FILES")
print(OUT_DIRECTORY)

# TODO add file check
# --------------------Extract case information from Info tab----------------


def processMetadata(caseInfo):
    try:
        print("\nLoading case metadata.\n")
        infoPD = pd.read_excel(
            caseInfo, sheet_name=clbInfoSheet, header=1, usecols="B,C"
        )
        if debug:
            print(infoPD)

        # Extract case metadata
        deviceType = infoPD.loc[infoPD["Name"] == "Device", ["Value"]].values[0]

        # Deal with rename of product
        # FIXME this is not a good way of dealing with this..use contains instead??
        try:
            paVersion = infoPD.loc[
                infoPD["Name"] == "Cellebrite Physical Analyzer version", ["Value"]
            ].values[0]
        except IndexError:
            paVersion = infoPD.loc[
                infoPD["Name"] == "UFED Physical Analyzer version", ["Value"]
            ].values[0]

        evidenceNumber = infoPD.loc[
            infoPD["Name"] == "Evidence number", ["Value"]
        ].values[0]

        caseNumber = infoPD.loc[infoPD["Name"] == "Case number", ["Value"]].values[0]

        extractionID = infoPD.loc[
            infoPD["Name"] == "      Extraction ID", ["Value"]
        ].values[0]

        logging.info("Loaded Metadata from: %s" % (str(caseInfo)))
        metadataFound = True

        # ---------------- Print Case metadata-----------------------------------
        print("Device Type: " + str(deviceType[0]))
        print("PA Version: " + str(paVersion[0]))
        print("Evidence Number: " + str(evidenceNumber[0]))
        print("Case Number: " + str(caseNumber[0]))
        print("Extraction ID: " + str(extractionID[0]))
        print(metadataFound)
        return metadataFound, str(caseNumber[0]), str(extractionID[0])

    # Value error is raised when the contacts sheet is not present.
    except ValueError:
        logging.info("Failed to load Metadata from: %s" % (str(caseInfo)))
        print("Unable to locate %s tab, is it present?" % (clbInfoSheet))
        print("Attempting to skip.")
        metadataFound = False
        return metadataFound, "NaN"
        pass


# Read in contacts data
def processContacts(contactData):
    metaResults = processMetadata(contactData)

    metaSuccess = metaResults[0]
    try:
        caseNumber = metaResults[1]
        extractionID = metaResults[2]
    except:
        pass
    print("Metadata extraction succesful: " + str(metaSuccess))

    # If metadata extraction was succesful continue
    if metaSuccess:

        exportFilename = contactData.split(".")[0]
        # print(infoPD)
        print(exportFilename)

        try:
            print("Loading contacts data")
            contactsPD = pd.read_excel(
                contactData, sheet_name=clbContactSheet, header=1, index_col="#"
            )
            contactsFound = True

        except ValueError:
            print("Unable to locate Contacts tab, is it present? \nSkipping.")
            contactsFound = False

        # Unstack contact data from nested cells
        if contactsFound:
            print(extractionID)
            print(caseNumber)
            contactsPD = contactsPD.drop("Entries", axis=1).join(
                contactsPD["Entries"]
                .str.split("\n", expand=True)
                .stack()
                .reset_index(level=1, drop=True)
                .rename("Entries")
            )
            # TODO add filenames in as a column
            # - Add in extraction GUID as a column
            contactsPD["InputFile"] = contactData
            # print(contactsPD["InputFile"])
            contactsPD["extractionGUID"] = extractionID
            contactsPD["caseNumber"] = caseNumber

            # TODO Remove  () and - from mobile and phone numbers
            # TODO Normalise Australian mobile numbers.
            # Split contact detail into contact types.
            contactsPD[contactTypeOutput] = contactsPD["Entries"].str.split(
                ":", expand=True
            )[0]
            contactsPD[contactOutput] = contactsPD["Entries"].str.split(
                ":", n=1, expand=True
            )[1]

            # Entries interest
            """
            Snapchat
            Phone-Mobile
            Phone-Home
            Instagram
            User ID-Facebook Id
            User ID-Username
            Twitter
            """

            # Create subset containing only phone numbers.
            phoneContactsPD = contactsPD.dropna(subset=[contactTypeOutput])
            phoneContactsPD = phoneContactsPD[
                phoneContactsPD[contactTypeOutput].str.contains(r"Phone")
            ]
            print(str(phoneContactsPD.shape[0]) + " phone numbers extracted.")
            phoneContactsPD = phoneContactsPD[List_of_Columns]
            phoneContactsPD[contactOutput] = phoneContactsPD[contactOutput].str.replace(
                "-", ""
            )

            # Dataframe of Telegram contacts
            try:
                telegramPD = contactsPD.dropna(subset=["Source"])
                telegramPD = telegramPD[telegramPD["Source"].str.contains(r"Telegram")]
                telegramPD.rename(columns={"Account": "sourceAccount"}, inplace=True)
                # Only export these columns
                telegramPD = telegramPD[
                    [
                        "extractionGUID",
                        "caseNumber",
                        "Name",
                        contactTypeOutput,
                        contactOutput,
                        "Interaction Statuses",
                        "Notes",
                        "Organizations",
                        "Addresses",
                        "User Tags",
                        "Device description",
                        "Source",
                        "sourceAccount",
                        "Deleted",
                        "Tag Note",
                        "Entries",
                    ]
                ]
                telegramProcess = True
                if debug:
                    print(telegramPD)
                print(str(telegramPD.shape[0]) + " Telegram contacts located")

            except Exception as ex:
                print(ex)
                print("Telegram module failed")
                telegramProcess = False

            # --------------------- Facebook ID's frame.---------------------------------
            try:
                facebookIDPD = contactsPD.dropna(subset=[contactTypeOutput])
                facebookIDPD = facebookIDPD[
                    facebookIDPD[contactTypeOutput].str.contains("User ID-Facebook Id")
                ]
                # There can be 2 entries, one for FB and one for FB messenger
                # Drop duplicate FB id's
                facebookIDPD = facebookIDPD.drop_duplicates(subset=["contactDetail"])
                facebookIDPD.rename(columns={"Account": "sourceAccount"}, inplace=True)
                facebookIDPD = facebookIDPD[
                    [
                        "Name",
                        contactTypeOutput,
                        contactOutput,
                        "Interaction Statuses",
                        "Notes",
                        "Source",
                        "sourceAccount",
                        "Entries",
                    ]
                ]
                print(
                    str(facebookIDPD.shape[0])
                    + " unique Facebook Account ID's extracted."
                )
                facebookProcess = True
            except Exception as ex:
                print(ex)
                print("Facebook moddule failed")
                facebookProcess = False

            # -------------Export to CSV files.------------------------------------------
            # FIXME Fix the output directory mess..
            print("Exporting CSV of Contacts")
            contactsPD.to_csv(
                os.path.join(OUT_DIRECTORY, exportFilename + "_contacts.csv")
            )
            print("Exporting Phone Numbers")
            phoneContactsPD.to_csv(
                os.path.join(OUT_DIRECTORY, exportFilename + "_phone.csv")
            )
            print("Exporting Facebook ID's")
            if facebookProcess:
                facebookIDPD.to_csv(
                    os.path.join(OUT_DIRECTORY, exportFilename + "_Facebook.csv")
                )

            if telegramProcess:
                print("Exporting Telegram")
                telegramPD.to_csv(
                    os.path.join(OUT_DIRECTORY, exportFilename + "_telegram.csv")
                )

            if debug:
                print(facebookIDPD)
    else:
        print("Skipping process due to error")
        # raise FormatError
        pass


# ----------------Run Process------------------------------------------------


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}".format(str(__author__), str(__version__)),
    )
    parser.add_argument(
        "-f", "--file", dest="inputFilename", help="Path to Excel spreadsheet."
    )

    args = parser.parse_args()

    if len(sys.argv) == 1:
        parser.print_help()
        parser.exit(1)

    if not os.path.exists(args.inputFilename):
        print("Error: '{}' does not exist or is not a file.".format(args.inputFilename))
        sys.exit(1)
processContacts(args.inputFilename)

"""
for y in inputFiles:
    processContacts(y)
"""
