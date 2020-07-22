# LiveWeather-Updated
[![Python Version](https://img.shields.io/badge/python-3.8.2-brightgreen.svg)](https://python.org)
[UPDATE] Now Fully Automated with error handling
A live weather automation script which exports, 'Live Data' from an weather API, in to an Excel Sheet for a regular interval. <br><br>
''' Used __openweathermap__ api '''

## Made with Python3 v3.8.2

Use the package manager [pip](https://pip.pypa.io/en/stable/) to install required libraries.


## Requirements
A "requirements.txt" is attached with the code.
Use it in cmd as 
```bash
pip install -r requirements.txt"
```
### Note
1. Change "weather.xlsx" file path in get_temperature() method.
2. Use keyboard interrupt (ctrl+c) to end the execution
3. Change time.sleep() value for changing the refresh interval. (Default = 1)
4. If shows "citynotfound" for correct city name, try changing "api_key"

## Contributing
Pull requests are welcome.<br> For major changes, please open an issue first to discuss what you would like to change.<br>
Please make sure to give proper credits in Readme.

## Credits
[Ayushman09](https://www.github.com/Ayushman09)
