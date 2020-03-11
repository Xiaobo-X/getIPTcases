
#import modules

# have to pip install requests first
# before requests can be imported to use
import requests

#re for regular expression
import re
import time
import os

#we will will urllib module to download pdf
import urllib
from bs4 import BeautifulSoup
import ast
import threading
import random
import xlsxwriter
import sys
from datetime import datetime
import re

#make a searchIPT class below
class searchIPT (object):
    #set request cookies, data form, etc
	#'rows' control number of search results to be returned
    #in each request
	
	#initialize
    def __init__ (self, searchTerms, instruction= '', rows='10', sort= 'decisionDateIndex_l desc'):
        self.searchTerms=searchTerms
        self.instruction= instruction
        self.rows= rows
        self.sort= sort
        self.url='https://forms.justice.govt.nz/solr/IPTV2/select?'
        self.facet= 'true'
        # if Search Terms has input value
        if self.instruction:
            self.q= 'policyprovision_s:"{0}" && ((abstract:{1} OR document_text_abstract:{1} OR decision:{1} OR document_text_uploaded:{1}))'.format(self.instruction,searchTerms)
        else:
            self.q= "((abstract:{0} OR document_text_abstract:{0} OR decision:{0} OR document_text_uploaded:{0}))".format(searchTerms)
        self.fq='jurisdictionCode_s:IPT AND categoryCode:RES'
        self.sort=sort
        self.jqueryWrf= 'jQuery112405626326185364463_1580959631345'
        self.queryString={'facet': self.facet,'rows': self.rows, 'fl': '*, score',\
                          'hl': 'true', 'hl.fl':'*',' hl.simple.pre': '<span class="highlight">',\
                          'hl.simple.post': '</span>', 'hl.fragsize': '10000',\
                          'hl.requireFieldMatch': 'true',  'hl.usePhraseHighlighter': 'true',\
                          'facet.limit': '-1', 'facet.mincount': '-1', 'sort': self.sort,\
                          'json.nl': 'map', 'q': self.q, 'fq': self.fq, 'wt': 'json','json.wrf': self.jqueryWrf,\
                          '_': '1580959631348'}
        self.requestHeaders= {'Accept': 'text/javascript, application/javascript, application/ecmascript, application/x-ecmascript, */*; q=0.01',\
                         'Accept-Encoding': 'gzip, deflate, br',\
                         'Accept-Language': 'en,zh-CN;q=0.9,zh;q=0.8,en-US;q=0.7,ja;q=0.6,zh-TW;q=0.5', \
                         'Connection': 'keep-alive',\
                         'Host': 'forms.justice.govt.nz',\
                         'Referer': 'https://forms.justice.govt.nz/search/IPT/Residence/',\
                         'Sec-Fetch-Mode': 'cors',\
                         'Sec-Fetch-Site': 'same-origin',\
                         'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36',\
                         'X-Requested-With': 'XMLHttpRequest'}
        # create a request Session
        self.session= requests.Session()

   # queryStringStart sets which page of search result
   # to start from
    def getSearchResult(self, queryStringStart=None):
        print ('start getSearchResult')
        #initialize
        self.queryStringStart= queryStringStart
        if queryStringStart != None:
            self.queryString['start'] = queryStringStart
        
        self.searchResponse= self.session.get (self.url, params=self.queryString, headers=self.requestHeaders)
        self.searchResponseText = str (self.searchResponse.text)
        #print ('self.searchResponseText: {}'.format (self.searchResponseText))
        self.searchResponseText = self.searchResponseText.replace (self.jqueryWrf+'(', '').replace(']}}})', ']}}}')
        #print ('self.searchResponseText: {}'.format (self.searchResponseText))

        #convert string dict to a dict 
        self.responseDict=  ast.literal_eval (self.searchResponseText)
        #print ('self.responseDict: {}'.format (self.responseDict))

        #count number of cases returned
        self.resultCount = int (self.responseDict['response']['numFound'])
        
        print ('self.resultCount: {}'.format( self.resultCount))
        self.responseDocs = self.responseDict['response']['docs']

        #wait for random time before make the next request - anti-anti-scraping
        timer = threading.Timer( (random.randint (1,1000) % 10 +1)/ 10 *1.389, self.doNothing)

        #return a list of typically 10 docs. one doc for one case
        return (self.responseDocs)

    def doNothing (self):
        pass

    def resultExcelList (self, responseCaseDict, interestedColumns):
        #get interestedColumns values from a dict in responseCaseDict
        # put them into rowList, and then put all rowList into rowsList
        self.interestedColumns= interestedColumns
        rowsList=[]
        for dict in responseCaseDict:
            rowList= []
            for column in self.interestedColumns:
                if column == 'abstractappealno_s':
                    value= str(dict['abstractappealno_s'][0])
                    if 'legacyIPTDecision_txt' in dict.keys() and str (dict['legacyIPTDecision_txt'][0]) == "Yes":
                        value= 'Residence Appeal ' + str (value) + str (' - ') + str (dict ['abstractapplicationdate_s'])
                    else:
                        value = re.sub(r"(.+?)\1+", r"\1", value)
                    rowList.append (value)
                else:
                    if column in dict.keys():
                        if 'legacyIPTDecision_txt' in dict.keys() and str (dict['legacyIPTDecision_txt'][0]) =='Yes' and  column=='abstractdateofdecision_s':
                            rowList.append (dict['abstractdecisiondate_s'][0])
                        else:
                            rowList.append (dict[column])
            rowsList.append (rowList)
        return rowsList

	#write rowList to local xls file
	# fieldsList is the xlsfile headers
    def writeToExcel (self, rowsList, fieldsList=[]):
        self.rowsList=rowsList
        self.fieldsList= fieldsList
        self.ExcelfileName='{0}- IPT -{1}.xlsx'.format (str (self.searchTerms) , datetime.now().strftime("%Y%m%d-%H%M%S"))
        workbook = xlsxwriter.Workbook(self.ExcelfileName)
        worksheet = workbook.add_worksheet()
        startRow,startColumn= 0,0
        if fieldsList != []:
            for field in fieldsList:
                worksheet.write(startRow, startColumn, field)
                startColumn+=1
        #write from startRow
        startrow=1
        for row in rowsList:
            startColumn=0
            for  cell in row:
                worksheet.write(startrow, startColumn, str (cell))
                startColumn+=1
            startrow +=1
        workbook.close()

searchIPT= searchIPT('false', instruction= 'A5.25' )
results= searchIPT.getSearchResult()
fields= [ 'abstractappealno_s', 'abstractdateofdecision_s', 'outcome_s','indexterms_s','policyprovision_s']
beginFrom= 350
resultCount= searchIPT.resultCount
while beginFrom <= int (resultCount):
    result = searchIPT.getSearchResult(queryStringStart=beginFrom)
    results.extend(result)
    beginFrom += int (searchIPT.rows) 
    
resultsExcelList = searchIPT.resultExcelList (results, fields )
searchIPT.writeToExcel(resultsExcelList, fieldsList= fields)









 



