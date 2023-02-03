"""
Extracts location data from a Cellebrite PA report and converts it to an ESRI friendly time format.

Data is Extracted from the Timeline tab of the Excel report.

Also has a feature to look for gaps in recording.

"""
import argparse
import logging
import pandas as pd
import os
import sys
from datetime import datetime, timedelta

# file = "input.xlsx"

# Details
__description__ = "Converts Cellebrite PA Garmin extracts to an ESRI compatible CSV.\n Loads data from the Timeline tab of the Excel Export"
__author__ = "facelessg00n"
__version__ = "0.1"

# Options
debug = False
findGaps = False
dateAfter = None
localConvert = True

localHour = 0.0
localMinute = 0.0

# ----------------------Config the logger------------------------------------
logging.basicConfig(
    filename="log.txt",
    format="%(levelname)s:%(asctime)s:%(message)s",
    level=logging.DEBUG,
)
# ---------------Functions --------------------------------------------------

# Setup function to load spreadsheet and columns of interest
def convertFile(inputFilename, dateAfter=None, gapFinder=None):
    if dateAfter is not None:
        print("Looking for dates after = " + str(dateAfter))
        dateCut = True
    else:
        dateCut = False

    print("Loading Excel file {}".format(inputFilename))
    logging.info("Loading Excel data from %s" % (str(inputFilename)))

    try:
        df = pd.read_excel(inputFilename, sheet_name="Timeline", header=1)
        df = df[["#", "Time", "Latitude", "Longitude"]]
    except Exception as e:
        print(e)
        exit()

    # Convert time format
    print("Converting Time Format")
    new = df["Time"].str.split("(", n=1, expand=True)
    df["DateTime"] = new[0]
    df["DateTime"] = pd.to_datetime(
        df["DateTime"], errors="raise", utc=True, format="%d/%m/%Y %I:%M:%S %p"
    )
    if debug == True:
        print(df.info())

    # Filter only data after this date
    if dateCut:
        try:
            df = df[(df["DateTime"] > dateAfter)]
        # "2020-02-01"
        except TypeError:
            print(
                "Type error has been raised, it is likely the input date format is incorect. This process will be skipped"
            )
            dateCut = False
            pass

    if localConvert:
        df["Local"] = df["DateTime"] + pd.Timedelta(
            hours=localHour, minutes=localMinute
        )

    # Find and report gaps in data recording.
    if gapFinder is not None:
        print(
            "Gap finder is looking for gaps of more than %s seconds." % (str(gapFinder))
        )
        gapData = True
    else:
        print("Gap finder is not looking for gaps in time")
        gapData = False

    if gapData:
        print("\nFinding gaps in time")
        df["GapFinder"] = df["DateTime"].diff().dt.seconds > gapFinder
        time_diff = df[df["GapFinder"] == True]
        print(str(time_diff.shape[0]) + " gaps in recording located.")
        gapData = True
        if debug:
            print(time_diff)

    # Export dataframes to CSV
    print("\nExporting CSV's")
    df.to_csv("locationData.csv", index=False, date_format="%Y/%m/%d %H:%M:%S")
    if gapData:
        time_diff.to_csv("gapData.csv", index=False, date_format="%Y/%m/%d %H:%M:%S")


# Command line input args
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}".format(str(__author__), str(__version__)),
    )

    parser.add_argument(
        "-f",
        "--file",
        dest="inputFilename",
        help="Path to input Excel Spreadsheet",
        # required=True,
    )

    parser.add_argument(
        "-g",
        "--gap",
        dest="gapSeconds",
        type=int,
        help="To detect gaps in time enter a time gap in seconds. 300 seconds is 5 minutes",
        default=None,
        required=False,
    )

    parser.add_argument(
        "-d",
        "--dateafter",
        dest="dateAfter",
        type=int,
        help="Filter only data after a certain date. Required format is YYYY-MM-DD. Useful for shrinking your dataset",
        required=False,
    )

    args = parser.parse_args()

    # display help message when no args are passed.
    if len(sys.argv) == 1:

        parser.print_help()
        sys.exit(1)

    # If no input show the help text.
    if not args.inputFilename:
        parser.print_help()
        parser.exit(1)

    # Check if the input file exists.
    if not os.path.exists(args.inputFilename):
        print("ERROR: '{}' does not exist or is not a file".format(args.inputFilename))
        sys.exit(1)

    if args.dateAfter is not None:
        dateAfter = args.dateAfter
        print("Date After Not none")
    else:
        dateAfter = None

    if args.gapSeconds is not None:
        gapSeconds = args.gapSeconds
        if debug:
            print("GapSeconds Not none")
    else:
        gapSeconds = None

    convertFile(args.inputFilename, gapFinder=gapSeconds, dateAfter=dateAfter)
