
"""

epXML2CSV

A utility to extract numbers from the tabular XML file created by EnergyPlus and put
them into a CSV format. The utility is especially useful when extracting data from
a large number of simulation results. The utility collects values from all XML files
in the directory or from a specific XML file.

  epXML2CSV csvfilename  extractfilename <xmlfilename>

If xmlfilename is provided, extract values from just that file. If the xmlfilename
is not specified, than etract values from all xml files in the directory.

A row is generated for the xmlfilename or for each xmlfile in the directory.
The columns correspond to each item extracted.

The extractfilename format is shown below and each row of the file
corresponds to one column of the resulting CSV file.

userHeading, reportNameString, forString, subReportString, nameString, headingString
userHeading, reportNameString, forString, subReportString, nameString, headingString
userHeading, reportNameString, forString, subReportString, nameString, headingString
...etc..

If the forString is * then new tables are generated at the end of the file for each
consisting of the one line for each file and one column for each possible forString
found across the files.

If the nameString is in the form #sum-string then the string is search for for all
all the values and take the sum of all records that include the string somewhere.

Copyright (c) 2015, Jason Glazer
All rights reserved.

Redistribution and use in source and binary forms, with or
without modification, are permitted provided that the
following conditions are met:

1. Redistributions of source code must retain the above
copyright notice, this list of conditions and the following
disclaimer.

2. Redistributions in binary form must reproduce the above 
copyright notice, this list of conditions and the following 
disclaimer in the documentation and/or other materials 
provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND
CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES,
INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT
NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

import xml.etree.cElementTree as ET
import sys
import os
import string
import csv
from pprint import pprint
from sets import Set

# Main function that returns a string based on the extract info
def getTextFromEPXML(root, reportS, forS, subRepS, nameS, headS):
    for repEl in root.findall('./' + reportS):
        if repEl.find('for').text == forS:
            #special type of #sum- extraction that sums value when matches are found
            if nameS[:5] == '#sum-':
                #print '>report: ', reportS
                #print '>for:    ', forS
                #print '>subRep: ', subRepS
                #print '>name:   ', nameS
                #print '>headS:  ', headS
                sum = 0.0
                searchString = nameS[5:].lower()
                #print "--searchString: ", searchString
                for subRepEl in repEl.findall('./' + subRepS):
                    #print subRepEl.tag
                    for child in subRepEl:
                        #print "  ",child.tag, child.text
                        if child.text.lower() == searchString:
                            try:
                                resultString = subRepEl.find(headS).text
                                #print resultString
                                if is_number(resultString):
                                    resultVal = float(resultString)
                                    sum = sum + resultVal
                            except AttributeError:
                                resultString = ''

                return str(sum)
            else:
                for subRepEl in repEl.findall('./' + subRepS):
                    # normal extraction of the value
                    if subRepEl.find('name').text == nameS:
                        try:
                            resultStringRaw = subRepEl.find(headS).text
                            resultString = resultStringRaw.replace(","," ").strip()
                        except:
                            resultString = 'NAME NOT FOUND'
                        if resultString != None:
                            return resultString
    return "~"

def is_number(s): # used by IL05
    try:
        float(s)
        return True
    except ValueError:
        return False

def epXML2CSV(argv=None):
    if argv == None:
        argv = sys.argv
    if len(argv) == 3:
        [progName, csvFileName, extractFileName] = argv
        listOfXmlFiles =[]
    elif len(argv) == 4:
        [progName, csvFileName, extractFileName,specXmlFileName] = argv
        listOfXmlFiles = [specXmlFileName,]
        if not os.path.exists(extractFileName):
            print "epXML2CSV Specified XML file file not found: ", specXmlFileName
            sys.exit(1)
    if not os.path.exists(extractFileName):
        print "epXML2CSV Extract file not found: ", extractFileName
        sys.exit(2)

    extractLines = []
    extractFile = open(extractFileName, 'r')
    extractReader = csv.reader(extractFile, delimiter=',')
    for row in extractReader:
        if row:
            if row[0][0] != "#":  #remove rows that start with comment symbol "#"
                extractLines.append(map(string.strip, row))

    extractFile.close()
    try:
        csvFile = open(csvFileName, 'wb')
    except IOError:
        print "epXML2CSV The csv output file appears to be open in another program: ", csvFileName
        sys.exit(3)
    csvWriter = csv.writer(csvFile, delimiter=',')

    firstHeadings = []
    listOfPulledUsingForWild = []
    userHeadingForWild = []

    # if no specific XML file was specified, then make list of all files in directory
    if len(listOfXmlFiles) == 0:
        for xmlFileName in os.listdir("."):
            if xmlFileName.endswith(".xml"):
                listOfXmlFiles.append(xmlFileName)

    for xmlFileName in listOfXmlFiles:
        # The following lists will contain the output for the CSV file, starting with name of XML file
        Headings = ["FileName"]
        Pulled = [xmlFileName]
        # the heavy lifting of parsing the XML file is done on the next line by the ElementTree library
        if not os.path.exists(xmlFileName):
            print "epXML2CSV File not found: ", xmlFileName
            continue
        tree = ET.parse(xmlFileName)
        root = tree.getroot()

        # -----------  Get the header information -------------
        #  <BuildingName>Building</BuildingName>
        #  <EnvironmentName>Chicago Ohare Intl Ap IL USA TMY3 WMO#=725300</EnvironmentName>
        #  <WeatherFileLocationTitle>Chicago Ohare Intl Ap IL USA TMY3 WMO#=725300</WeatherFileLocationTitle>
        #  <ProgramVersion>EnergyPlus-Windows-32 8.0.0.008, YMD=2013.05.16 11:07</ProgramVersion>
        #  <SimulationTimestamp>
        #    <Date>
        #      2013-05-16
        #    </Date>
        #    <Time>
        #      11:07:53
        #    </Time>
        #  </SimulationTimestamp>

        Headings.append('BuildingName')
        Pulled.append(root.find('BuildingName').text)

        Headings.append('EnvironmentName')
        Pulled.append(root.find('EnvironmentName').text)

        Headings.append('WeatherFileLocationTitle')
        Pulled.append(root.find('WeatherFileLocationTitle').text)

        Headings.append('ProgramVersion')
        Pulled.append(root.find('ProgramVersion').text.replace(","," "))

        Headings.append('SimulationTimestamp')
        Pulled.append(root.find('SimulationTimestamp').find('Date').text.strip() + ' ' + root.find('SimulationTimestamp').find('Time').text.strip())

        for [userHeadingStr, reportStr, forStr, subReportStr, nameStr, headingStr] in extractLines:
            #when using a wildcard on the FOR an extra column is needed one for the name of each FOR
            if forStr == '*':
                #loop through all instances of the report for each "for"
                forFounds = []
                PulledFromFor = {}
                for repEl in root.findall('./' + reportStr):
                    forFounds.append(repEl.find('for').text)
                for forFound in forFounds:
                    PulledFromFor[forFound] = getTextFromEPXML(root, reportStr, forFound, subReportStr, nameStr, headingStr)
                listOfPulledUsingForWild.append([xmlFileName,userHeadingStr,PulledFromFor])
                if userHeadingStr not in userHeadingForWild:
                    userHeadingForWild.append(userHeadingStr)
                Pulled.append('SEE BELOW') # a filler to make sure a space is taken
            else:
                Pulled.append(getTextFromEPXML(root, reportStr, forStr, subReportStr, nameStr, headingStr))
            Headings.append(userHeadingStr)

        if Headings != firstHeadings:
            firstHeadings = Headings
            csvWriter.writerow(Headings)

        csvWriter.writerow(Pulled)

        print "File: " + xmlFileName + "   Number of values: " + str(len(Pulled))
 #   pprint(listOfPulledUsingForWild,stream=csvFile)
    for userHeading in userHeadingForWild:
        #create a table for each user defined heading with columns of FOR's and rows of the filenames
        csvFile.write('\n') #leave a blank line
        csvFile.write('\n') #leave a blank line
        csvFile.write(userHeading + '\n')
        prevKeys = []
        for [file,userHeadingInList,resDict] in listOfPulledUsingForWild:
            if userHeading == userHeadingInList:
                curKeys = resDict.keys()
                if curKeys != prevKeys:  #only print a new list of
                    csvFile.write('\n') #leave a blank line
                    row = ['File',]
                    row.extend(curKeys)
                    csvWriter.writerow(row)
                    prevKeys = curKeys
                row = []
                row.append(file)
                for curKey in curKeys:
                    row.append(resDict[curKey])
                csvWriter.writerow(row)
    csvFile.close()

if __name__ == "__main__":
    sys.exit(epXML2CSV())
