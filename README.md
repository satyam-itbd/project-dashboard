# Excel Live Web Page

This project starts a small Python web page that reads `data.xlsx` every time the page is loaded.

## 1. Install dependencies

```powershell
& "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\python.exe" -m pip install -r requirements.txt
```

## 2. Create or edit the Excel file

Create a file named `data.xlsx` in this folder.

Example columns:

| Name | Department | Score |
|------|------------|-------|
| Asha | Sales      | 91    |
| Ravi | Support    | 88    |

## 3. Start the web app

```powershell
& "$env:LOCALAPPDATA\Python\pythoncore-3.14-64\python.exe" app.py
```

## 4. Open the page

Visit:

```text
http://127.0.0.1:5000
```

When you change and save `data.xlsx`, refresh the browser page to see the new values.
