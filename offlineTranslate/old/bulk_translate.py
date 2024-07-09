# Bulk Translation of Axiom formatted Excels containing messages
# Made in South Australia
# Unapologetically formatted with Black
#
#
# Changelog
#
# v0.1 Initial Concept

import argparse
import json
import pandas as pd
import requests
import os
import sys

# ----------------- Settings live here ------------------------

__description__ = "Utilises a Libretranslate server to translate messages from Axiom formatted Excel spreadsheets. Messages are loaded from a column titled 'Messages.'"
__author__ = "facelessg00n"
__version__ = "0.1"

banner = """
 ██████  ███████ ███████ ██      ██ ███    ██ ███████     ████████ ██████   █████  ███    ██ ███████ ██       █████  ████████ ███████ 
██    ██ ██      ██      ██      ██ ████   ██ ██             ██    ██   ██ ██   ██ ████   ██ ██      ██      ██   ██    ██    ██      
██    ██ █████   █████   ██      ██ ██ ██  ██ █████          ██    ██████  ███████ ██ ██  ██ ███████ ██      ███████    ██    █████   
██    ██ ██      ██      ██      ██ ██  ██ ██ ██             ██    ██   ██ ██   ██ ██  ██ ██      ██ ██      ██   ██    ██    ██      
 ██████  ██      ██      ███████ ██ ██   ████ ███████        ██    ██   ██ ██   ██ ██   ████ ███████ ███████ ██   ██    ██    ███████ 
                                                                                                                                      
                                                                                                                                      """

# Debug mode, will print errors etc
debug = False

serverURL = "http://localhost:5000"
# Endpoints
#           /translate - translation
#           /languages - supported languages

# Name of the column where the messages to be translated are found.
# This can be modified to suit other Excel column names if desired
inputColumn = "Message"


# Check is server is reachable and able to process a request.
def serverCheck():
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

    # FIXME - Handle connection errors, can probably be done better.
    except ConnectionRefusedError:
        print(
            f"Server connection refused - {serverURL}, is the address correct? \n\nExiting"
        )
        sys.exit()
    except Exception as e:
        print(f"Unable to connect, ERROR: {e}")
        sys.exit()


# Loads Excel into dataframe and translates messages
def loadAndTranslate(inputFile, inputLanguage):
    # Check we can hit the server before we start
    serverCheck()
    head, tail = os.path.split(inputFile)
    fileName = tail.split(".")[0]
    # Load Excel into Dataframe "df" and check for messages column.
    df = pd.read_excel(inputFile)

    if inputColumn not in df.columns:
        print("Required message column not found")
        sys.exit(1)

    # Load Messages Column to list and print some stats
    messages_nan_count = df["Message"].isna().sum()
    messages = df["Message"].tolist()
    print(f"{len(messages)} messages")
    print(f"{messages_nan_count} blank rows")

    results = []
    loopCount = 0
    for message in messages:
        # If no language code is specified use Auto Translate
        if inputLanguage == None:
            translated_text = translate_text(message, None)
        # Else manual translation
        else:
            translated_text = translate_text(message, inputLanguage)

        if debug:
            print(translated_text)
        results.append(translated_text)
        print(f"Processing message {loopCount} of {len(messages)}")
        loopCount = loopCount + 1

        # ------------- Write backup file every 100 messages ----------------------------------------
        if len(results) % 100 == 0:
            print("Writing backup")
            backup_frame = pd.DataFrame(results)

            try:
                backup_frame.to_excel(
                    f"{fileName}_backup.xlsx",
                    index=False,
                    columns=[
                        "detected_language",
                        "detected_confidence",
                        "success",
                        "input",
                        "translatedText",
                    ],
                )
            except:
                print("Writing Excel Bakcup failed")
                pass

            try:
                backup_frame.to_csv(
                    f"{fileName}_backup.csv",
                    encoding="utf-16",
                    columns=[
                        "detected_language",
                        "detected_confidence",
                        "success",
                        "input",
                        "translatedText",
                    ],
                )
            except:
                print("Writing CSV backup failed")
                pass

    # ------------------ Write output file -----------------------------------------------------------------
    print("Translation complete - Writing file")
    outputFrame = pd.DataFrame(results)

    try:
        outputFrame.to_excel(
            f"{fileName}_translated.xlsx",
            index=False,
            columns=[
                "detected_language",
                "detected_confidence",
                "success",
                "input",
                "translatedText",
            ],
        )
    except:
        print("Writing Excel file")
        pass

    try:
        outputFrame.to_csv(
            f"{fileName}_translated.csv",
            encoding="utf-16",
            columns=[
                "detected_language",
                "detected_confidence",
                "success",
                "input",
                "translatedText",
            ],
        )
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
            print("Manual Lanugage Detection {}".format(inputLang))
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
        print("Blank row found, skipping")
        output = {
            "detected_language": None,
            "detected_confidence": None,
            "translatedText": None,
            "success": False,
        }
        output["input"] = inputText
        return output

    else:
        headers = {"Content-Type": "application/json"}
        response = requests.post(
            f"{serverURL}/translate", data=payload, headers=headers
        )
        if response.status_code == 200:
            results = response.json()
            if debug:
                print(f"{inputText} and {response.json()}")
            try:
                answer = results
                # Server response style is different for Auto or Manual language selection
                if inputLang is not None:
                    output = {
                        "detected_language": f"Manual - {inputLang}",
                        "detected_confidence": None,
                        "translatedText": answer.get("translatedText"),
                        "success": True,
                    }
                else:
                    output = {
                        "detected_language": results.get("detectedLanguage")[
                            "language"
                        ],
                        "detected_confidence": results.get("detectedLanguage")[
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
                "detected_language": None,
                "detected_confidence": None,
                "translatedText": None,
                "success": f"Error: {response.status_code, results.get}",
            }
            output["input"] = inputText
            return output


# Retrieve list of allowed languages from the server
def getLanguages(printVals):
    AllowedLangs = []
    supportedLanguages = requests.get(f"{serverURL}/languages").json()
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
    serverCheck()
    print(f"Checking server {serverURL} for supported languages")
    try:
        supportedLanguages = getLanguages(False)
        if len(supportedLanguages) == 0:
            print("Supported Languages not found")
            supportedLanguages = ["0"]
        else:
            print(f"Languages found - {supportedLanguages} \n\n")

    except Exception as e:
        print(e)

    parser = argparse.ArgumentParser(
        description=__description__,
        epilog="Developed by {}, version {}".format(str(__author__), str(__version__)),
    )

    parser.add_argument(
        "-f", "--file", dest="inputFilePath", help="Path to Axiom formatted excel file"
    )
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
        "-g",
        "--getlangs",
        dest="getLangs",
        action="store_true",
        help="Get supported language codes and names from server",
        required=False,
        default=True,
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
        loadAndTranslate(args.inputFilePath, None)

    if args.inputFilePath and args.inputLanguage:
        if not os.path.exists(args.inputFilePath):
            print(
                "ERROR: {} does not exist or is not a file".format(args.inputFilePath)
            )
            sys.exit(1)
        print(f"Input language set to {args.inputLanguage}")
        loadAndTranslate(args.inputFilePath, args.inputLanguage)

    if args.getLangs:
        getLanguages(True)
