import requests   # pip install requests
import httplib2   # pip install httplib2
from bs4 import BeautifulSoup, SoupStrainer   # pip install beautifulsoup4
import warnings

##############################
# CHANGE HERE
# bLatestOnly = True >> Only want the latest one
# bLatestOnly = False >> All available versions
bLatestOnly = False
##############################

# clean output :-)
warnings.filterwarnings("ignore")

# URL to SCRAP
scrapURL    = 'https://docs.microsoft.com/sv-se/officeupdates/update-history-microsoft365-apps-by-date'

# list to save stuff to
availableVersions = list()

http = httplib2.Http()
status, response = http.request(scrapURL)

# parse the response
soup = BeautifulSoup(response,"html.parser")

# find all tables
tables = soup.find_all('table')

# get all headers
strBuff = ''

# read the headers from table 2
if len(tables) == 2:   # check so we find 2 tables, and if yes, check table 2
    for thead in tables[1].find_all('thead'):   # headers are in thead
        for th in thead.find_all('th'):
            if len(strBuff) == 0:
                strBuff = str(th.string)
            else:
                strBuff = strBuff + ";" + str(th.string)
else:
    exit()

arrHeader = strBuff.split(';')

availableVersions.append([arrHeader[0],arrHeader[1],'Version','Build'])   # save the first row, the headers, to the availableVersions list

# data variables used in next for section
strYear = ''
strReleaseDate = ''
strSemiAnnualEntChnnl = ''
iRows = 0

for tbody in tables[1].find_all('tbody'):   # find tbody
    for tr in tbody.find_all('tr'):   # for each row, scrap data
        x = -1
        for td in tr.find_all('td'):   # data is in the td
            x = x + 1

            if x == 0:   # year can be found in column 1
                if str(td.string) != 'None':
                    strYear = str(td.string)
            elif  x == 1:   # relase date can be found in column 2 (Microsoft is not consistent for the older ones, so therefor a special else to take care of that)
                if str(td.string) != 'None':
                    strReleaseDate = str(td.string)
                elif str(td).find('br') > 0:
                    strTemp = str(td).replace('<td style=\"text-align: left;\">','').replace('<br/></td>','')
                    strTemp = strTemp.strip()
                    strReleaseDate = strTemp
            elif  x == 5:   # Semi-Annual Enterprise Channel version can be found in column 6
                tempsoup = BeautifulSoup(str(td),"html.parser")
                iii = len(tempsoup.find_all('a'))
                if iii > 0:
                    for href in tempsoup.find_all('a'):
                        if len(strSemiAnnualEntChnnl) == 0:
                            strSemiAnnualEntChnnl = str(href.string)
                        else:
                            strSemiAnnualEntChnnl = strSemiAnnualEntChnnl + ';' + str(href.string)
                    strData = strSemiAnnualEntChnnl.split(';')
                    for d in strData:   # for each version found, add it to the availableVersions list
                        strVer = d.split(' (Build ')[0].replace('Version ','')   # get version
                        strBuild = d.split(' (Build ')[1].replace(')','')   # get build number
                        print(strYear + ';' + strReleaseDate + ';' + strVer + ';' + strBuild)
                        availableVersions.append([strYear,strReleaseDate,strVer,strBuild])
                        iRows = iRows + 1

                    strSemiAnnualEntChnnl = ''   # clean variable

                    if iRows > 0 and bLatestOnly:   # if we only want the latest version, exit now
                        break
        if iRows > 0 and bLatestOnly:   # if we only want the latest version, exit now
            break
    if iRows > 0 and bLatestOnly:   # if we only want the latest version, exit now
        break

# write everything to a txt file (csv separated)
f = open("ScrapTheWeb_v2.txt", "wt")
for elem in availableVersions:
    f.write(elem[0] + ';' + elem[1] + ';' + elem[2] + ';' + elem[3] + '\n')
f.close()    
