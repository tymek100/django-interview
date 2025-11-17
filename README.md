# DJANGO interview task

## Dependencies

### Python
- For Linux / Mac it should be preinstalled.
- For Windows you can download it here: https://www.python.org/downloads/windows/

### PDM
Mac / Linux:
```
curl -sSL https://pdm-project.org/install-pdm.py | python3 -
```
Windows:
```
powershell -ExecutionPolicy ByPass -c "irm https://pdm-project.org/install-pdm.py | py -"
```

## Run the app
```
pdm run python manage.py migrate
pdm run python manage.py runserver
```

### Call the API
```
curl -X POST "http://127.0.0.1:8000/api/excel-summary/" \
  -H "Accept: application/json" \
  -F "file=@example.xlsx" \
  -F "columns=CURRENT USD" \
  -F "columns=CURRENT CAD"
```
