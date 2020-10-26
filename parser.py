from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
import lxml.html
from optparse import OptionParser
import openpyxl
import os
from datetime import datetime
from random import randint
from pprint import pprint
from string import ascii_uppercase
from pdb import set_trace


driver=webdriver.Chrome()
driver.maximize_window()

workbook=openpyxl.Workbook()


betsTypes=['home-draw-away', 'double-chance', 'both-teams-to-score', 'clean-sheet', 'half-time-full-time', 'odd-even']
betTypes=['Both Teams To Score', 'Clean Sheet', 'Double Chance', 'Half Time / Full Time', 'Home Draw Away', 'Odd / Even']
bookmakers=['William Hill', 'Bwin', 'Betago', '188bet', 'Pinnacle', 'Boyle Sports']


def processArgs():
    parser = OptionParser()
    parser.add_option("-n", "--num", dest="number",
                      help="Number of matches to parse", default=20)
    options,args=parser.parse_args()
    return options


def getMainListing():
    driver.get('https://betbrain.ru')
    #nextMatches
    sleep(randint(3,5))
    try:
        sportsTabs=driver.find_elements_by_xpath('//a[@class="SportsBarTab"]')
        ActionChains(driver).move_to_element(sportsTabs[0]).click(sportsTabs[0]).perform()
    except IndexError as e:
        try:
            sleep(randint(3,5))
            sportsTabs=driver.find_elements_by_xpath('//a[@class="SportsBarTab"]')
            assert sportsTabs
            ActionChains(driver).move_to_element(sportsTabs[0]).click(sportsTabs[0]).perform()
        except:
            for _ in range(5):
                sleep(randint(3, 5))
                sportsTabs = driver.find_elements_by_xpath('//a[@class="SportsBarTab"]')
                if sportsTabs: break
            ActionChains(driver).move_to_element(sportsTabs[0]).click(sportsTabs[0]).perform()

    #football
    sleep(randint(3,5))
    try:
        sportsIcons=driver.find_elements_by_xpath('//a[@class="SportsIcon"]')
        ActionChains(driver).move_to_element(sportsIcons[0]).click(sportsIcons[0]).perform()
    except IndexError as e:
        try:
            sleep(randint(3,5))
            sportsIcons = driver.find_elements_by_xpath('//a[@class="SportsIcon"]')
            assert sportsIcons
            ActionChains(driver).move_to_element(sportsIcons[0]).click(sportsIcons[0]).perform()
        except:
            for _ in range(5):
                sleep(randint(3, 5))
                sportsIcons = driver.find_elements_by_xpath('//a[@class="SportsIcon"]')
                if sportsIcons: break
            ActionChains(driver).move_to_element(sportsIcons[0]).click(sportsIcons[0]).perform()


def getMatches(toLoad=20):

    #loadMore
    while toLoad > 20:
        loadButton=driver.find_element_by_xpath('//button[@class="Button SportsBoxAll LoadMore"]')
        ActionChains(driver).move_to_element(loadButton).click(loadButton).perform()
        sleep(randint(3,5))
        toLoad-=20

    #findMatchElements
    try:
        matches = driver.find_elements_by_xpath('//li[@class="Match"]')
        matches[7]
    except IndexError as e:
        try:
            sleep(randint(3, 5))
            matches = driver.find_elements_by_xpath('//li[@class="Match"]')
            matches[10]
        except:
            for _ in range(5):
                sleep(randint(3, 5))
                matches = driver.find_elements_by_xpath('//li[@class="Match"]')
                if matches: break

    #matchName/matchUrl
    matchItems=list()
    for matchItem in matches:
        match = matchItem.find_element_by_css_selector('.MatchTitleLink')
        matchName = match.text.split('\n')[0]
        matchURL = match.get_attribute('href')
        matchItems.append((matchName,matchURL))

    return matchItems[:toLoad]



try:
    siteBaseURL='https://betbrain.ru'

    #parseCLIOpts
    opts=processArgs()
    toLoad=int(opts.number)

    #getListingView
    getMainListing()

    #20-first matchesDefault
    sleep(randint(3, 5))
    matchItems=getMatches(toLoad)

    #countIterCycles
    cycles=int(toLoad/4)
    if toLoad % 4!=0:
        cycles+=1

    #processBy4tabs
    for _ in range(cycles):
        matches=matchItems[:4]

        #minNumOfTabs
        width=min(len(matchItems),4)

        # matchesTabsPreopened
        for matchItem in matches:
            driver.execute_script(f'window.open("{matchItem[1]}");')
            sleep(randint(3,6))

        # matchesTabs
        for i,matchItem in enumerate(matches):
            driver.switch_to.window(driver.window_handles[width-i])
            output=list()
            j=3
            sleep(randint(3,6))

            #newSheet
            matchSheet=workbook.create_sheet(matchItem[0])
            print(matchItem[0])

            #neededBetTypes
            for betType in betsTypes:
                result=list()

                #betTypeUrl
                currentUrl = driver.current_url.split('/')
                currentUrl[-3] = betType
                driver.get('/'.join(currentUrl))

                #lxml
                sleep(randint(3,9))
                try:
                    page=lxml.html.fromstring(driver.page_source)
                    table=page.xpath('//div[@class="OddsTable StaticOddsTable"]')[0]

                except IndexError as e:
                    try:
                        sleep(randint(3,9))
                        page = lxml.html.fromstring(driver.page_source)
                        table = page.xpath('//div[@class="OddsTable StaticOddsTable"]')[0]
                    except:
                        try:
                            for _ in range(5):
                                sleep(randint(3,9))
                                page = lxml.html.fromstring(driver.page_source)
                                table = page.xpath('//div[@class="OddsTable StaticOddsTable"]')
                                if table: break
                            table=table[0]

                        #noTableAtAll
                        except: continue

                #rf/other OddsTables
                oddsTables=table.getchildren()

                if len(oddsTables)==0: continue
                idsRF, idsOther=list(), list()

                #betHeaders
                header=['']
                try:
                    otHeader=oddsTables[0].xpath('.//ul[@class="OTHead OTRow"]')[0]
                except IndexError as e:
                    otHeader = oddsTables[1].xpath('.//ul[@class="OTHead OTRow"]')[0]
                for otMark in otHeader.getchildren():
                    header.append(otMark.text_content())

                #oddsTables
                for ot in oddsTables:
                    try:
                        #bookie'sData
                        bookmakersRF=ot.xpath('.//ol[@class="OTBookmakersContainer"]')[0].getchildren()
                        for i,bookie in enumerate(bookmakersRF):
                            bookieName=bookie.xpath('.//div/a/span/span')[0].text_content()
                            bookieRow=[bookieName]
                            if bookieName in bookmakers:
                                j+=1
                                idsRF.append(i)
                                result.append(bookieRow)

                        #otData
                        oData=ot.xpath('.//div[@class="OTOddsData"]')[0].getchildren()
                        oDataG=tableG.find_elements_by_xpath('.//div[@class="OTOddsData"]/div')

                        for k,i in enumerate(idsRF):

                            oRow=oData[i].xpath('.//ul[@class="OTRow"]')[0]

                            for col in oRow.xpath('.//li/a/span/span'):
                                result[k].append(col.text_content())

                            if len(result[k])>len(header):
                                result[k]=[result[k][0]]+result[k][1-len(header):]

                    #noData=>nextTable
                    except Exception as e:
                        print(e)
                        continue

                #insertHeaders
                result.insert(0,[betType])
                result.insert(1,header)

                #addEmptyNewLine
                result.append([' '])

                output.append(result)

            #fillEXELSheet
            # set_trace()
            # print(output)
            j=0
            for betRow in output:
                for row in betRow:
                    #skipEmptyHeaders
                    if row == ['']*len(row): continue
                    j+=1
                    for y,cell in enumerate(row):
                        matchSheet[f'{ascii_uppercase[y]}{j}'] = cell
            driver.close()
        driver.switch_to.window(driver.window_handles[0])
        matchItems=matchItems[4:]

    #match

except Exception as e:
    print(e)
finally:
    #keepPreviousSpreadSheet
    try:
        os.rename(f'matches-{toLoad}.xlsx',f'matches-{toLoad}-previous.xlsx')
    except: pass
    driver.quit()
    os.system('pkill -f "/snap/chromium/1328/usr/lib/chromium-browser/chrome"')
    os.system('pkill -f "/snap/chromium/1328/usr/lib/chromium-browser/chromedriver"')
    #rmDefaultSheet
    try:
        workbook.remove(workbook['Sheet'])
    except: pass
    workbook.save(f'matches-{toLoad}.xlsx')
