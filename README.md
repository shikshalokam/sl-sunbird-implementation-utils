# Program Update Automation Script

This Python script automates updating programs by calling the Programs Update API using
program IDs provided in a CSV file.

## How to Run

```bash
pip install requests
python3 main.py program_ids.csv
```

## Files
- update_programs.py : Main script
- config.ini : API config and tokens
- program_ids.csv : Input CSV with program IDs
- api_responses.txt : Output file
- install the dependencies   pip install -r requirements.txt

