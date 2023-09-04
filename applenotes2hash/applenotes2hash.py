# Extracts password protected hashes from Apple Notes
#
# Elements from script from Dhiru Kholia <kholia at kth.se>
# https://github.com/openwall/john/blob/bleeding-jumbo/run/applenotes2john.py
#
# Formatted with Black
#
# ------Changes----------
#
# V0.1 - Initial release
#

import argparse
import binascii
import glob
import os
import sys
import sqlite3
import shutil
import zipfile

PY3 = sys.version_info[0] == 3

if not PY3:
    reload(sys)
    sys.setdefaultencoding("utf8")

__description__ = "Extracts and converts Apple Note hashes to Hashcat and JTR format"
__author__ = "facelessg00n"
__version__ = "0.1"

formatType = []
notesFile = "NoteStore.sqlite"
targetPath = os.getcwd() + "/temp"
debug = False

# ------------- Functions live here -----------------------------------


def makeTempFolder():
    try:
        # print("Creating temporary folder")
        os.makedirs(targetPath)
    except OSError as e:
        # print(e)
        # print("Temporary folder exists")
        # print("Purging directory")
        shutil.rmtree(targetPath)
        try:
            # print("Creating temporary folder")
            os.makedirs(targetPath)
        except:
            # print("Something has gone horribly wrong")
            exit()


# Check it is a zip file and extract relevant file
def checkZip(z):
    if zipfile.is_zipfile(z):
        # print("This is a Zip File")
        with zipfile.ZipFile(z) as file:
            zippedFiles = file.namelist()
            filePath = [x for x in zippedFiles if x.endswith(notesFile)]
            if debug:
                print("Located file at path : {}".format(filePath))
                print("Extracting to temp file")
            file.extract(filePath[0], targetPath)

    else:
        print("this does not appear to be a zip file")


def processGrayShift(x, formatType):
    formatType = formatType
    try:
        makeTempFolder()
    except Exception as e:
        print(e)
    checkZip(x)
    inputFile = glob.glob("./**/NoteStore.sqlite", recursive=True)
    if debug:
        print("Using" + str(inputFile[0]) + " as the input file for Cache.")
    extractHash(inputFile[0], formatType)


# Functionality below lifted from
# https://github.com/openwall/john/blob/bleeding-jumbo/run/applenotes2john.py


def extractHash(inputFile, formatType):
    db = sqlite3.connect(inputFile)
    cursor = db.cursor()
    rows = cursor.execute(
        "SELECT Z_PK, ZCRYPTOITERATIONCOUNT, ZCRYPTOSALT, ZCRYPTOWRAPPEDKEY, ZPASSWORDHINT, ZCRYPTOVERIFIER, ZISPASSWORDPROTECTED FROM ZICCLOUDSYNCINGOBJECT"
    )
    for row in rows:
        iden, iterations, salt, fhash, hint, shash, is_protected = row
        if fhash is None:
            phash = shash
        else:
            phash = fhash
        if hint is None:
            hint = "None"
            # NOTE: is_protected can be zero even if iterations value is non-zero!
            # This was tested on macOS 10.13.2 with cloud syncing turned off.
        if iterations == 0:  # is this a safer check than checking is_protected?
            continue
        if phash is None:
            continue
        phash = binascii.hexlify(phash)
        salt = binascii.hexlify(salt)
        if PY3:
            phash = str(phash, "ascii")
            salt = str(salt, "ascii")
        fname = os.path.basename(inputFile)
        # For John
        if formatType == "JOHN":
            sys.stdout.write(
                "%s:$ASN$*%d*%d*%s*%s:::::%s\n"
                % (fname, iden, iterations, salt, phash, hint)
            )
        # For Hashcat
        elif formatType == "HASHCAT":
            sys.stdout.write("$ASN$*%d*%d*%s*%s\n" % (iden, iterations, salt, phash))

        else:
            print("Invalid or no format type set")
        db.close


# ----------- Argument Parser ---------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}, version {}".format(str(__author__), str(__version__)),
    )

    parser.add_argument(
        "-f", "--file", dest="notesFile", help="Path to NoteStore.sqlite"
    )
    parser.add_argument(
        "-g", "--grayshift", dest="grayshiftINPUT", help="Path to Grayshift Extract"
    )
    parser.add_argument(
        "-t",
        "--type",
        dest="formatType",
        help="Output format type, JOHN or HASHCAT, defaults to JOHN. Hashcat Mode is 16200",
        choices=["HASHCAT", "JOHN"],
        default="JOHN",
        required=False,
    )

    args = parser.parse_args()
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

if args.notesFile:
    if not os.path.exists(args.notesFile):
        print("ERROR: {} does not exist or is not a file".format(args.notesFile))
        sys.exit(1)
    extractHash(notesFile, args.formatType)

if args.grayshiftINPUT:
    if not os.path.exists(args.grayshiftINPUT):
        print("ERROR: {} does not exist or is not a file".format(args.grayshiftINPUT))
        sys.exit(1)
    processGrayShift(args.grayshiftINPUT, args.formatType)
