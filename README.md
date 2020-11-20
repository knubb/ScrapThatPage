# ScrapThatPage
## What
This python script Scraps a Microsoft webpageL to get all releases in the Semi-Annual Enterprise Channel and what date it was released:  
https://docs.microsoft.com/sv-se/officeupdates/update-history-microsoft365-apps-by-date

## Why
Since I needed that information :-) ...and with this nifty little script I can set-it-and-forget-it! #loveautomation

## How
I used beautifulsoup4 to scrap the page and retrieve information needed.

### Prerequisites
Install following library’s using pip (needed to run script)  
* pip install requests  
* pip install httplib2  
* pip install beautifulsoup4  

### Run the script
Just run it with following command:  
```
python ScrapThatPage.py
```

### Output
The output CSV will be saved in the directory where the script is executed.
