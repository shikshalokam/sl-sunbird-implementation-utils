import csv
import requests
import json
import configparser
import sys
import os

if len(sys.argv) != 2:
    print("Usage: python3 main.py <input_csv_file>")
    sys.exit(1)

CSV_FILE = sys.argv[1]

if not os.path.exists(CSV_FILE):
    print(f"Error: File not found -> {CSV_FILE}")
    sys.exit(1)

OUTPUT_FILE = "api_responses.txt"

config = configparser.ConfigParser()
config.read("config.ini")

BASE_URL = config["API"]["base_url"]


# check headers key and value 
HEADERS = {
    "X-authenticated-user-token": config["API"]["x_user_token"],
    "internal-access-token": config["API"]["internal_access_token"],
    "Authorization": config["API"]["authorization"],
    "Content-Type": "application/json"
}

BODY = {
    "status": "inactive",
    "isDeleted": True
}

with open(OUTPUT_FILE, "w") as out_file:
    with open(CSV_FILE, "r") as csv_file:
        reader = csv.DictReader(csv_file)

        if "id" not in reader.fieldnames:
            print("Error: CSV must contain 'id' column")
            sys.exit(1)

        for row in reader:
            program_id = row["id"].strip()
            url = f"{BASE_URL}/{program_id}"

            try:
                response = requests.post(url, headers=HEADERS, json=BODY)

                out_file.write(f"==== Program ID: {program_id} ====\n")
                out_file.write(f"Status Code: {response.status_code}\n")

                try:
                    out_file.write(json.dumps(response.json(), indent=2))
                except Exception:
                    out_file.write(response.text)

                out_file.write("\n\n")
                print(f"Processed: {program_id}")

            except Exception as e:
                out_file.write(f"==== Program ID: {program_id} ====\n")
                out_file.write(f"ERROR: {str(e)}\n\n")

print("Execution completed. Check api_responses.txt")
