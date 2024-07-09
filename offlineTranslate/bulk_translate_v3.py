# Bulk Translation of Axiom formatted Excels containing messages
# Made in South Australia
# Unapologetically formatted with Black
#
# Changelog
# v0.3 Handle network errors... oops
# v0.2 Change to output full content of the input sheet
#      Handle Cellebrite and Axiom files
# v0.1 Initial Concept

import argparse
import json
import pandas as pd
import requests
import os
import sys
from tqdm import tqdm
from time import sleep

# ----------------- Settings live here ------------------------

__description__ = "Utilises a Libretranslate server to translate messages from Excel spreadsheets. By default messages are loaded from a column titled 'Message'."
__author__ = "facelessg00n"
__version__ = "0.3"

banner = """
 ██████  ███████ ███████ ██      ██ ███    ██ ███████     ████████ ██████   █████  ███    ██ ███████ ██       █████  ████████ ███████ 
██    ██ ██      ██      ██      ██ ████   ██ ██             ██    ██   ██ ██   ██ ████   ██ ██      ██      ██   ██    ██    ██      
██    ██ █████   █████   ██      ██ ██ ██  ██ █████          ██    ██████  ███████ ██ ██  ██ ███████ ██      ███████    ██    █████   
██    ██ ██      ██      ██      ██ ██  ██ ██ ██             ██    ██   ██ ██   ██ ██  ██ ██      ██ ██      ██   ██    ██    ██      
 ██████  ██      ██      ███████ ██ ██   ████ ███████        ██    ██   ██ ██   ██ ██   ████ ███████ ███████ ██   ██    ██    ███████ 
                                                                                                                                      
                                                                                                                                      """

# Debug mode, will print errors etc
debug = False

# if being compiled with a GUI
# Keeps window alive if connection fails
hasGUI = True

serverURL = "http://localhost:5000"
CONNECTION_TIMEOUT = 3
RESPONSE_TIMEOUT = 60
#
# Endpoints
#           /translate - translation
#           /languages - supported languages

# Name of the column where the messages to be translated are found.
# This can be modified to suit other Excel column names if desired
inputColumn = "Message"
inputSheets = ["Chats", "Instant Messages"]
sheetName = "Chats"
headerRow = 1

translationColumns = [
    "detectedLanguage",
    "detectedConfidence",
    "success",
    "input",
    "translatedText",
]


# Check is server is reachable and able to process a request.
def serverCheck(serverURL):
    print(f"Testing we can reach server {serverURL}")
    headers = {"Content-Type": "application/json"}
    payload = json.dumps(
        {
            "q": "Buenos días señor",
            "source": "auto",
            "target": "en",
            "format": "text",
            "api_key": None,
        }
    )
    try:
        response = requests.post(
            f"{serverURL}/translate", data=payload, headers=headers
        )
        if response.status_code == 404:
            print("ERROR: 404, server not found, check server address.")
            sys.exit(1)
        elif response.status_code == 400:
            print("ERROR: Invalid request sent - exiting")
            sys.exit(1)
        elif response.status_code == 200:
            print("Server located, testing translation")
            print(response.json())
            return "SERVER_OK"

    # FIXME - Handle connection errors, can probably be done better.
    except ConnectionRefusedError:
        print(
            f"Server connection refused - {serverURL}, is the address correct? \n\nExiting"
        )
        if not hasGUI:
            sys.exit()
    except Exception as e:
        print(f"Unable to connect, ERROR: {e}")
        if not hasGUI:
            sys.exit()


# Loads Excel into dataframe and translates messages
def loadAndTranslate(inputFile, inputLanguage, inputSheet, isCellebrite):
    # Check we can hit the server before we start
    serverCheck(serverURL)
    head, tail = os.path.split(inputFile)
    fileName = tail.split(".")[0]

    if isCellebrite:
        inputHeader = 1
        inputColumn = "Body"
    else:
        inputHeader = 0
        inputColumn = "Message"

    # Load Excel into Dataframe "df" and check for messages column.
    if inputSheet:
        print("There is an input sheet")
        df = pd.read_excel(inputFile, sheet_name=inputSheet, header=inputHeader)
    else:
        print("There is no input sheet specified")
        df = pd.read_excel(inputFile, header=inputHeader)

    if debug:
        df = df.head(25)

    if inputColumn not in df.columns:
        print("Required message column not found, is this a Cellbrite Formatted Excel?")
        sys.exit(1)

    # Load Messages Column to list and print some stats
    messages_nan_count = df[inputColumn].isna().sum()
    messages = df[inputColumn].tolist()
    print(f"{len(messages)} messages")
    print(f"{messages_nan_count} blank rows")

    results = []
    loopCount = 1
    for message in tqdm(messages, desc="Translating messages", ascii="░▒█"):
        # If no language code is specified use Auto Translate
        if inputLanguage == None:
            translated_text = translate_text(message, None)
        # Else manual translation
        else:
            translated_text = translate_text(message, inputLanguage)

        if debug:
            print(translated_text)
        results.append(translated_text)
        tqdm.write(f"Processing message {loopCount} of {len(messages)}")
        # print(f"Processing message {loopCount} of {len(messages)}")
        loopCount = loopCount + 1

        # ------------- Write backup file every 100 messages ----------------------------------------
        if len(results) % 100 == 0:
            tqdm.write("Writing backup")
            backup_frame = pd.DataFrame(results)

            try:
                backup_frame.to_csv(
                    f"{fileName}_backup.csv",
                    encoding="utf-16",
                    columns=translationColumns,
                )
            except:
                print("Writing CSV backup failed")
                pass

    # ------------------ Write output file -----------------------------------------------------------------
    print("Translation complete - Writing file")
    # Get colum positon to insert new data
    bodyPosition = df.columns.get_loc(inputColumn) + 1
    # Splitting orig frame into to then concat with new data
    df1_part1 = df.iloc[:, :bodyPosition]
    df1_part2 = df.iloc[:, bodyPosition:]
    outputFrame = pd.concat([df1_part1, pd.DataFrame(results), df1_part2], axis=1)

    try:
        outputFrame.to_excel(f"{fileName}_translated.xlsx", index=False)
    except:
        print("Writing Excel failed")
        pass

    try:
        outputFrame.to_csv(f"{fileName}_translated.csv", encoding="utf-16")
    except:
        print("Writing CSV failed")
        pass

    print("Process complete - Exiting.")


# ------------------ Translates text with selected language -----------------------------------------------
def translate_text(inputText, inputLang, api_key=None):
    # For future implementation
    if api_key is not None:
        API_KEY = api_key
    else:
        API_KEY = None

    if inputLang is not None:
        if debug:
            print("Manual Lanugage Selection {}".format(inputLang))
        payload = json.dumps(
            {
                "q": inputText,
                "source": inputLang,
                "target": "en",
                "format": "text",
                "api_key": API_KEY,
            }
        )
    else:
        if debug:
            print("Auto language detection enabled".format(inputLang))
        payload = json.dumps(
            {
                "q": inputText,
                "source": "auto",
                "target": "en",
                "format": "text",
                "api_key": API_KEY,
            }
        )

    # Detect blank rows and skip to prevent error being thrown by server / speeds up process
    if inputText == None or pd.isna(inputText):
        tqdm.write("Blank row found, skipping")
        output = {
            "detectedLanguage": None,
            "detectedConfidence": None,
            "translatedText": None,
            "success": False,
        }
        output["input"] = inputText
        return output

    # If row is not blank, attempt to translate it
    else:
        headers = {"Content-Type": "application/json"}
        try:
            # Max Attempt for retries
            MAX_ATTEMPTS = 5

            response = requests.post(
                f"{serverURL}/translate",
                data=payload,
                headers=headers,
                timeout=(CONNECTION_TIMEOUT, RESPONSE_TIMEOUT),
            )

        # Handle a read timeout error, sleep 2 seconds then try again
        except requests.ReadTimeout:

            while MAX_ATTEMPTS > 0:
                try:
                    tqdm.write("Read Timeout error, retrying")
                    sleep(2)
                    response = requests.post(
                        f"{serverURL}/translate",
                        data=payload,
                        headers=headers,
                    )
                    output["input"] = inputText
                    return output

                except Exception:
                    MAX_ATTEMPTS -= 1
                    continue
            else:
                output = {
                    "detectedLanguage": None,
                    "detectedConfidence": None,
                    "translatedText": None,
                    "success": "False: Error: Read Timeout ",
                }
                output["input"] = inputText
                return output

        # Handle a connection dropout, sleep 2 seconds and try again
        except requests.ConnectionError:
            while MAX_ATTEMPTS > 0:
                try:
                    tqdm.write("Connection Error - Retrying")
                    sleep(2)
                    response = requests.post(
                        f"{serverURL}/translate", data=payload, headers=headers
                    )
                    output["input"] = inputText
                    return output

                except Exception:
                    MAX_ATTEMPTS -= 1
                    continue
            else:
                print("Failed")
                output = {
                    "detectedLanguage": None,
                    "detectedConfidence": None,
                    "translatedText": None,
                    "success": "False: Error: Connection Error",
                }
                output["input"] = inputText
                return output

        except Exception as e:
            tqdm.write(f"Unhandled exception {e}")
            output = {
                "detectedLanguage": None,
                "detectedConfidence": None,
                "translatedText": None,
                "success": f"False: Error: {e}",
            }
            output["input"] = inputText
            return output

        if response.status_code == 200:
            results = response.json()
            if debug:
                print(f"{inputText} and {response.json()}")
            try:
                answer = results
                # Server response style is different for Auto or Manual language selection
                if inputLang is not None:
                    output = {
                        "detectedLanguage": f"Manual - {inputLang}",
                        "detectedConfidence": None,
                        "translatedText": answer.get("translatedText"),
                        "success": True,
                    }
                else:
                    output = {
                        "detectedLanguage": results.get("detectedLanguage")["language"],
                        "detectedConfidence": results.get("detectedLanguage")[
                            "confidence"
                        ],
                        "translatedText": answer.get("translatedText"),
                        "success": True,
                    }

                output["input"] = inputText
                return output
            except Exception as e:
                print(e)

        elif response.status_code == 400:
            print("Invalid request")
            output = {
                "detectedLanguage": None,
                "detectedConfidence": None,
                "translatedText": None,
                "success": f"Error: {response.status_code, results.get}",
            }
            output["input"] = inputText
            return output


# Retrieve list of alowed languages from the server
def getLanguages(printVals):
    AllowedLangs = []
    try:
        supportedLanguages = requests.get(f"{serverURL}/languages").json()
    except:
        print("Supported Languages not found")
        supportedLanguages = []
        pass

    for langItem in supportedLanguages:
        if printVals:
            print(
                f"Language Code: {langItem['code']} Language Name: {langItem['name']}"
            )
        AllowedLangs.append(langItem["code"])
    return AllowedLangs


# ---------------------------- Argument Parser ------------------------

if __name__ == "__main__":
    print(banner)
    if debug:
        print("WARNING DEBUG MODE IS ACTIVE")
    serverCheck(serverURL)
    print(f"Checking server {serverURL} for supported languages")
    try:
        supportedLanguages = getLanguages(False)
        if len(supportedLanguages) == 0:
            print("Supported Languages not found")
            supportedLanguages = []
        else:
            print(f"Languages found - {supportedLanguages} \n\n")

    except Exception as e:
        print(e)

    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}, version {}".format(str(__author__), str(__version__)),
    )

    parser.add_argument("-f", "--file", dest="inputFilePath", help="Path to Excel File")
    parser.add_argument(
        "-s",
        "--server",
        dest="translationServer",
        help="Address of translation server if not localhost or hardcoded",
        required=False,
    )

    parser.add_argument(
        "-l",
        "--language",
        dest="inputLanguage",
        help="Language code for input text - optional but can greatly improve accuracy",
        required=False,
        choices=supportedLanguages,
    )

    parser.add_argument(
        "-e",
        "--excelSheet",
        dest="inputSheet",
        help="Sheet name within Excel file to be translated",
        required=False,
        choices=inputSheets,
    )

    parser.add_argument(
        "-c",
        "--isCellebrite",
        dest="isCellebrite",
        help="If file originated from Cellebrite, header starts at 1, and message column is called 'Body'",
        required=False,
        action="store_true",
        default=False,
    )

    parser.add_argument(
        "-g",
        "--getlangs",
        dest="getLangs",
        action="store_true",
        help="Get supported language codes and names from server",
        required=False,
        default=False,
    )

    args = parser.parse_args()
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(1)

    if args.inputFilePath and not args.inputLanguage:
        if not os.path.exists(args.inputFilePath):
            print(
                "ERROR: {} does not exist or is not a file".format(args.inputFilePath)
            )
            sys.exit(1)
        loadAndTranslate(args.inputFilePath, None, args.inputSheet, args.isCellebrite)

    if args.inputFilePath and args.inputLanguage:
        if not os.path.exists(args.inputFilePath):
            print(
                "ERROR: {} does not exist or is not a file".format(args.inputFilePath)
            )
            sys.exit(1)
        print(f"Input language set to {args.inputLanguage}")
        loadAndTranslate(
            args.inputFilePath, args.inputLanguage, args.inputSheet, args.isCellebrite
        )

    if args.getLangs:
        getLanguages(True)
