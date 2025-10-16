(Get-Content comparison_app.py -Raw).Split([Environment]::NewLine) | Select-Object -Index (109..159)
