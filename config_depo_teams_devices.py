'''
Copyright (c) 2020 Cisco and/or its affiliates.

This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at

               https://developer.cisco.com/docs/licenses

All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.

'''

import pygsheets
from webexteamssdk import WebexTeamsAPI
import requests, os, json
from time import strftime, localtime
import datetime
import base64
from multiprocessing.pool import ThreadPool
from config import SPREADSHEET_NAME, WEBEX_SITE
from dotenv import load_dotenv
load_dotenv()

DEVICES_LOCAL_INT_ACCOUNT=os.environ["DEVICES_LOCAL_INT_ACCOUNT"]
DEVICES_LOCAL_INT_PWD=os.environ["DEVICES_LOCAL_INT_PWD"]

class DepMember:
    def __init__(self, name, email, deviceMAC):
        self.name = name
        self.email = email
        self.isHost = False
        self.deviceMAC = deviceMAC

    def set_host(self):
        self.isHost = True

    def set_deviceMAC(self, deviceMAC):
        self.deviceMAC=deviceMAC

class DepTeam:
    def __init__(self, name):
        self.name = name
        self.members = []    # creates a new empty list of members for each team
        self.start = ""
        self.end = ""
        self.host = ""
        self.pmr_sip_uri = ""
        self.host_macro_and_panel=""
        self.participants_macro_and_panel = ""

    def add_member(self, member):
        self.members.append(member)
    def set_host(self, theHost):
        theHost.set_host()
        self.host = theHost
    def set_start(self, theStart):
        self.start=theStart
    def set_end(self, theEnd):
        self.end=theEnd
    def set_pmr_sip_uri(self, thePMRSIPURI):
        self.pmr_sip_uri=thePMRSIPURI
    def set_host_macro_and_panel(self, thehost_macro_and_panel):
        self.host_macro_and_panel=thehost_macro_and_panel
    def set_participants_macro_and_panel(self, theparticipants_macro_and_panel):
        self.participants_macro_and_panel=theparticipants_macro_and_panel

api = WebexTeamsAPI(os.environ["WEBEX_TEAMS_ACCESS_TOKEN"])
gc = pygsheets.authorize(client_secret='./client_id.json')

#read the macro/panel templates file into a variable for both Host and Participants
theHostXMLTemplate=open("TeamHostMacroPanel.xml", "r").read()
thePartXMLTemplate=open("TeamPartMacroPanel.xml", "r").read()
thePanelRowXMLTemplate=open("TeamPanelRows.xml", "r").read()
theBlankXMLTemplate=open("TeamBlankMacroPanel.xml", "r").read()
theMacroEnableXML=open("macros-enable.xml", "r").read()


#retrieve using requests since not covered in WebexTeamsSDK. Only 100 at a time
url = "https://webexapis.com/v1/devices?max=100"
payload = {}
headers = {
  'Authorization': 'Bearer '+ os.environ["WEBEX_TEAMS_ACCESS_TOKEN"]
}

try:
    response = requests.request("GET", url, headers=headers, data = payload)
    theHeaders=response.headers

    theDevices=json.loads(response.text)
    theHeaders=response.headers
    #check for additional pages since we only requested the first 100 devices
    while 'Link' in theHeaders:
        #iterate to get the next 100
        iter_url=response.links
        #check to see if the link in question is for a next page
        if 'next' in iter_url:
            #yes, go ahead and navigate to the next URL and call the API again
            next_url=iter_url['next']['url']
            print("Going get next page at: ",next_url)
            response = requests.request("GET", next_url, headers=headers, data=payload)
            nextDevices=json.loads(response.text)
            #add the next page of devices to theDevices
            theDevices["items"].extend(nextDevices["items"])
            theHeaders = response.headers
except requests.exceptions.HTTPError as errh:
    print ("Http Error:",errh)
except requests.exceptions.ConnectionError as errc:
    print ("Error Connecting:",errc)
except requests.exceptions.Timeout as errt:
    print ("Timeout Error:",errt)
except requests.exceptions.RequestException as err:
    print ("OOps: Something Else",err)

if 'errors' in theDevices:
    print("Error: ",theDevices['errors'][0]['description'])
    quit()

#thePersonalDevices is a dictionary of all personal devices that are in ControlHub
#the key for thePersonalDevices dict is the MAC address of the device
thePersonalDevices={}

#theUserIndex is a Dict that contains the MAC address of the device associated to a user based on their
#email address.
theUserIndex={}

# theUsersFullObjects is a dict indexed by email address of all person objects obtained from Webex
# to use for updating the display name later
theUsersFullObjects={}

#theTeams is a dict indext by team name of all the Teams that are being configured
theTeams={}

#For all devices that have a PersonID (personal devices) extract the email address of the owner
#and add it to the devices structure
for theDevice in theDevices["items"]:
    if 'personId' in theDevice:
        theDisplayName=theDevice['displayName']
        theMAC=theDevice['mac']
        thePersonId=theDevice['personId']
        print(f'The device {theDisplayName} with MAC {theMAC} has a personId of {thePersonId}. Retrieving person email...')
        thePerson=api.people.get(thePersonId)
        theEmail=thePerson.emails[0]
        theOwnerDisplayName=thePerson.displayName
        theOwnerLicenses=thePerson.licenses
        thePersonalDevices[theMAC]={
                "id":theDevice['id'],
                "product":theDevice['product'],
                "serial":theDevice['serial'],
                "connectionStatus":theDevice['connectionStatus'],
                "ownerEmail":theEmail,
                "ownerDisplayName": theOwnerDisplayName,
                "ownerLicenses": theOwnerLicenses,
                "inSheet": False,
                "sendMacros": False,
                "sendClear": False,
                "macro" : "",
                "ip":theDevice['ip'],
                "sheetRowNum": 0
        }
        #now let's update theUserIndex for quick checking which device belongs to a user by email
        theUserIndex[theEmail]=theMAC
        #we are also keeping all of the person objects so that later we can update the display name
        theUsersFullObjects[theEmail]=thePerson
        print("The person email for this device is: ",theEmail)


# Open the Google Sheets Spreadsheet
sh = gc.open(SPREADSHEET_NAME)

# Select the DATA workseet where we will read input and
# report status.
wks = sh.worksheet_by_title("DATA")

# Select legend worksheet to be able to set status by copying from
# appropriate cells
legend_wks = sh.worksheet_by_title("Legend")

# set up references to the various status errors described in the
# 'Legend' worksheet
leg_mac_pushed='A2'
leg_dis_pushed='A3'
leg_unreachable='A4'
leg_autherror='A5'
leg_missing_user='A6'
leg_missing_device='A7'
leg_missing_device_user='A8'
leg_device_missmatch='A9'
leg_nodep_unreachable='A10'
leg_nodep_autherror='A11'
theGreyColor = (0.9372549, 0.9372549, 0.9372549, 0)

def setStatusCell(theRow,theLegendCell):
    #print("skipping update of status and stamp for speed....")
    global wks, legend_wks
    rStatusToSet = wks.cell('H' + str(theRow))
    rStatusToUse = legend_wks.cell(theLegendCell)
    rStatusToSet.value = rStatusToUse.value
    rStatusToSet.color = rStatusToUse.color
    rStatusToSet.update()
    rTimeStampToSet = wks.cell('I' + str(theRow))
    rTimeStampToSet.value=strftime("%m/%d/%Y %H:%M:%S", localtime())
    rTimeStampToSet.color=theGreyColor
    rTimeStampToSet.update()

# Count the number of rows in the worksheet that have data
rcount=0
for row in wks:
    rcount += 1
print("The worksheet has these many rows: ",rcount)

# check for existing rows in the worksheet, mark any discrepancies
# with the correct status code and add any missing devices to the end
# in new rows

#firt set rnum to 1 in case the Google sheet has no data rows because we refrence rnum below as the next empty row
rnum=1

for rnum in range(2,rcount+1):
    print("Checking row: ",rnum)
    rEmail=wks.cell('A'+str(rnum)).value
    theGreyColor=(0.9372549, 0.9372549, 0.9372549, 0)
    rMac=wks.cell('B'+str(rnum)).value
    #first check for matching email with MAC to see if we can mark the device as present
    #in the sheet and if there is something in the DepName field then mark it for pushing macros
    if (rMac in thePersonalDevices) and (rEmail in theUserIndex):
        #keep track of the Sheet row number to be able to report status when iterating throught devices
        #we will send macros to at the end
        thePersonalDevices[rMac]['sheetRowNum']=rnum
        if thePersonalDevices[rMac]['ownerEmail']==rEmail:
            #the right user and MAC are in sheet, mark it as such
            thePersonalDevices[rMac]['inSheet'] = True
            print("This device is correctly in the sheet!: ",rMac)
            rDepName=wks.cell('D'+str(rnum)).value
            if rDepName!='':
                # there is a value in the DepName field, we will be sending macros to the device
                # and updating the display name using object from theUsersFullObjects
                thePersonalDevices[rMac]['sendMacros']=True

                # collect the rest of the data from the row. We already have rEmail, rMac and rDepName from above
                rDisplayName=wks.cell('C'+str(rnum)).value
                # keep a separate variable for the final value of the DisplayName to use in controlhub
                theUpdatedDisplayName=rDisplayName

                rIsHost=wks.cell('E'+str(rnum)).value

                print("rIsHost is of type: ", type(rIsHost))

                if rIsHost=="TRUE":
                    rDepStart = wks.cell('F' + str(rnum)).value
                    rDepEnd = wks.cell('G' + str(rnum)).value
                    # append the string (Host) to the DisplayName if not already there in the sheet
                    if theUpdatedDisplayName[-6:]!="(Host)":
                        theUpdatedDisplayName=theUpdatedDisplayName+"(Host)"

                if rDepName=="CLEAR":
                    # we need to assign the macro to send as the empty one
                    # and setting the display name using object from theUsersFullObjects
                    # back to a generic username
                    thePersonalDevices[rMac]['sendClear'] = True
                    thePersonalDevices[rMac]['macro'] = theBlankXMLTemplate
                    theDispNameCell=wks.cell('C' + str(rnum))

                    # since we are clearing this entry, we want to set the display name for the user in
                    # controlhub to just the username
                    theUpdatedDisplayName=rEmail.split("@")[0]
                else:
                    # now we need to set up a macro template for this device which will
                    # later be filled out with all the details specific to the teams it
                    # belongs to. We cannot just generate the macro at this point because
                    # we have not yet iterated through the whole sheet so we do not have the
                    # complete list of members for the team nor have we validated that it is
                    # correct and with just 1 host
                    thePersonalDevices[rMac]['macro'] = "the specific macro template"
                    if rDepName not in theTeams:
                        theTeams[rDepName]=DepTeam(rDepName)
                    aMember = DepMember(theUpdatedDisplayName, rEmail, rMac)
                    theTeams[rDepName].add_member(aMember)
                    if rIsHost == "TRUE":
                        # this row represents the host of the team, so we will configure deposition Start and
                        # end dates that should be in the respective fields in the same row and also will
                        # mark the team member as host and set the Personal Meeting Room (PMR) URI for the
                        # team to match that of the host user so it can be programmed in the macro as the
                        # destination to call to join the deposition meetings.
                        theTeams[rDepName].set_host(aMember)
                        #TODO: Validate that rDepStart<=rDepEnd before assigning
                        theTeams[rDepName].set_start(rDepStart)
                        theTeams[rDepName].set_end(rDepEnd)
                        theTeams[rDepName].set_pmr_sip_uri(rEmail.split("@")[0]+"@"+WEBEX_SITE)

            # Update the displayname in the sheet and in controlhub whenever we are to take action
            # on a row (CLEAR or set correct macros). We have already set the correct value in the
            # theUpdatedDisplayName variable
            # First update the sheet
            theDispNameCell = wks.cell('C' + str(rnum))
            theDispNameCell.value = theUpdatedDisplayName
            theDispNameCell.update()
            # Then in control hub, where we have to first extract the entire people object and send
            # and updated one
            thePersObjToUpdate=theUsersFullObjects[rEmail]
            api.people.update(thePersObjToUpdate.id,
                              emails=thePersObjToUpdate.emails,
                              displayName=theUpdatedDisplayName,
                              firstName=thePersObjToUpdate.firstName,
                              lastName=thePersObjToUpdate.lastName,
                              avatar=thePersObjToUpdate.avatar,
                              orgId=thePersObjToUpdate.orgId,
                              roles=thePersObjToUpdate.roles,
                              licenses=thePersObjToUpdate.licenses
                              )
            # let's start checking for users and their licenses to match up hosts
            print("Finished pre-processing ",theUpdatedDisplayName," marked as isHost=",rIsHost," with licenses: ",thePersObjToUpdate.licenses)


        else:
            # there is a valid email and MAC in the row, but they do not match as per controlhub!
            # need to set the right status for that row... the right user/device will be added at the end
            # of the worksheet if no other row has the right match
            setStatusCell(rnum,leg_device_missmatch)
    else:
        #one of the two, email or MAC, from the sheet, does not match ControlHub personal devices/users
        if (rMac not in thePersonalDevices) and (rEmail not in theUserIndex):
            # the user and MAC are both missing from control hub as related to personal devices!
            setStatusCell(rnum, leg_missing_device_user)
        elif rEmail not in theUserIndex:
            #the user is does not have a personal device in control hub
            setStatusCell(rnum, leg_missing_user)
        else:
            #the device MAC is not in ControlHub as a personal device
            setStatusCell(rnum, leg_missing_device)



# now we are ready to add new rows for devices that were not in the sheet before. We add the MAC and owner
# email which are not supposed to be edited, but as a convenience we add the initial owner display name
# that can be changed.
# Remember that rnum contains the next empty row in the sheet after looping from existing rows above
for theMACKey in thePersonalDevices:
    theDevice=thePersonalDevices[theMACKey]
    print("Checking device: ",theMACKey)
    if not theDevice['inSheet']:
        rnum += 1
        cellEmail = wks.cell('A' + str(rnum))
        cellEmail.value=theDevice['ownerEmail']
        cellEmail.color=theGreyColor
        cellEmail.update()
        cellMac = wks.cell('B' + str(rnum))
        cellMac.value=theMACKey
        cellMac.color=theGreyColor
        cellMac.update()
        cellDN = wks.cell('C' + str(rnum))
        cellDN.value=theDevice['ownerDisplayName']
        cellDN.update()
        theDevice['inSheet']=True

#Now that the sheet has been checked for discrepancies and new entries have been added
#the following must be added:

#TODO validate that depositions are complete with a HOST and team members, need to define what to do with
# incomplete ones.

for key, aTeam in theTeams.items():
    print("Processing team: ",key)
    #TODO: validate that the team has a host
    #TODO: validate there is at least one more member in the team besides the host
    #TODO: validate START and END dates

    #need to change date format from MM/DD/YYYY to YYYY-MM-DD for the macro to process correctly in JS
    theStartTime = datetime.datetime.strptime(aTeam.start, "%m/%d/%Y").strftime("%Y-%m-%d")
    theEndTime=datetime.datetime.strptime(aTeam.end, "%m/%d/%Y").strftime("%Y-%m-%d")

    # FIRST: we process the Participants XML templates to change the Team PMR URI and start and end dates
    thePartXML=thePartXMLTemplate.replace("TEAMPMRURI",aTeam.pmr_sip_uri)
    thePartXML=thePartXML.replace("STARTSTUB",theStartTime)
    thePartXML=thePartXML.replace("ENDSTUB", theEndTime)

    # afterwards,  we do the same for the Host XML Template.....
    theHostXML=theHostXMLTemplate.replace("TEAMPMRURI",aTeam.pmr_sip_uri)
    theHostXML=theHostXML.replace("STARTSTUB",theStartTime)
    theHostXML=theHostXML.replace("ENDSTUB", theEndTime)

    # but we are not done with the template for the host, we need to iterate through all
    # participants to add the rows for the panel to call each one
    # while at it, we can set the thePartXML string to each participants object
    # since that one is ready to be pushed out to their devices

    # initialize variables useful to build the host macro XML
    theMemberEmailListStr=""
    thePanelRows=""
    theIndex=0
    for theMember in aTeam.members:
        if not theMember.isHost:
            #first set the macro for the participant
            thePersonalDevices[theMember.deviceMAC]['macro'] = thePartXML
            #now extract the info you will use to finish substitutions in theHostXML
            # we start with building up the list of email adresses (user URIs) to dial
            if theMemberEmailListStr=="":
                theMemberEmailListStr="'"+theMember.email+"'"
            else:
                theMemberEmailListStr=theMemberEmailListStr+", '"+theMember.email+"'"
            #now we insert one row in the custom call panel per each non-host user
            aRow=thePanelRowXMLTemplate
            aRow=aRow.replace("DISPLAYNAMESTUB",theMember.name)
            aRow=aRow.replace("INDEXSTUB",str(theIndex))
            theIndex+=1
            thePanelRows=thePanelRows+aRow

    # after we iterated through all members, we can replace MEMBERSURILISTSTUB in the XML string
    theHostXML=theHostXML.replace("MEMBERSURILISTSTUB", theMemberEmailListStr)
    # we can also insert all the rows in the PANELROWSTUB placeholder
    theHostXML = theHostXML.replace("PANELROWSTUB", thePanelRows)
    # finally, set the macro for the host of this team
    thePersonalDevices[aTeam.host.deviceMAC]['macro'] = theHostXML

#TODO: need to figure out how to overwrite existing panels)

theDestinationDevices=[]
#check which devices need to send macros to add to theDestinationDevices[]
#and write out to a log file for now to verify all macros:
Logfile = "Generated_macros_report.txt"
with open(Logfile, "a+") as text_file:
    for theMACKey in thePersonalDevices:
        theDevice = thePersonalDevices[theMACKey]
        if theDevice['sendMacros']:
            print("Writing out macros for device: ", theMACKey)
            text_file.write( "Macros for device: "+theMACKey+ '\n')
            text_file.write(theDevice['macro'])
            text_file.write('\n\n\n')
            #put the entire dict that includes 'ip' and 'macro' keys into theDestinationDevices array to know
            #where to send macros and which ones
            theDestinationDevices.append(theDevice)

# This function is where the magic happens. I am using request to open the xml file and then posting the content
# to the url of each TP endpoint based on the IP address obtained from the CSV file
# NB that http needs to be enabled on the TP endpoint otherwise you will get a 302 error.
def do_upload(aDestDevice):

        try:
            Logfile="Macros_send_report.txt"
            payload=theMacroEnableXML
            url = "http://{}/putxml".format(aDestDevice['ip'])
            userpass = DEVICES_LOCAL_INT_ACCOUNT + ':' + DEVICES_LOCAL_INT_PWD
            encoded_u = base64.b64encode(userpass.encode()).decode()
            headers = {
                'Content-Type': 'text/xml',
                'Authorization': 'Basic '+encoded_u,
                'Content-Type': 'text/plain'
            }
            print('-'*40)
            print('Enabling Macros on {}'.format(aDestDevice['ip']))
            response = requests.request("POST", url, headers=headers, data=payload, verify=False)
            print(response.text)

            with open(Logfile, "a+") as text_file:
                text_file.write("\nThe Status of Macros enabling on codec IP {} |---->>>".format(aDestDevice['ip']) + '\n')
                text_file.write(response.text)

            payload2=aDestDevice['macro']
            print('-' * 40)
            print('Configuring In-Room Control and Macros on {}'.format(aDestDevice['ip']))

            response = requests.request("POST", url, headers=headers, data=payload2, verify=False)
            print(response.text)
            with open(Logfile, "a+") as text_file:
                text_file.write("The Status of In-Room Control and Macros Config on codec {} |---->>".format(aDestDevice['ip']) + '\n')
                text_file.write(response.text)
            # set status in corresponding row for success if you reach here!
            if aDestDevice['sendClear']:
                setStatusCell(aDestDevice['sheetRowNum'], leg_dis_pushed)
            else:
                setStatusCell(aDestDevice['sheetRowNum'], leg_mac_pushed)

        except requests.exceptions.HTTPError as errh:
            print("Http Error:", errh)
            text_file.write('Http error talking to {} : {}'.format(aDestDevice['ip'], errh))
            # set status in corresponding Sheet row for failure
            if aDestDevice['sendClear']:
                setStatusCell(aDestDevice['sheetRowNum'], leg_nodep_autherror)
            else:
                setStatusCell(aDestDevice['sheetRowNum'], leg_autherror)
        except requests.exceptions.ConnectionError as errc:
            print("Error Connecting:", errc)
            text_file.write('failed to connect to {} : {}'.format(aDestDevice['ip'], errc))
            # set status in corresponding Sheet row for failure
            if aDestDevice['sendClear']:
                setStatusCell(aDestDevice['sheetRowNum'], leg_nodep_unreachable)
            else:
                setStatusCell(aDestDevice['sheetRowNum'], leg_unreachable)
        except requests.exceptions.Timeout as errt:
            print("Timeout Error:", errt)
            text_file.write('timeout connecting to {} : {}'.format(aDestDevice['ip'], errt))
            # set status in corresponding Sheet row for failure
            if aDestDevice['sendClear']:
                setStatusCell(aDestDevice['sheetRowNum'], leg_nodep_unreachable)
            else:
                setStatusCell(aDestDevice['sheetRowNum'], leg_unreachable)

        except requests.exceptions.RequestException as err:
            print("Other Error: ", err)
            text_file.write('Other error trying to send to {} : {}'.format(aDestDevice['ip'], err))
            # set status in corresponding Sheet row for failure
            if aDestDevice['sendClear']:
                setStatusCell(aDestDevice['sheetRowNum'], leg_nodep_unreachable)
            else:
                setStatusCell(aDestDevice['sheetRowNum'], leg_unreachable)

def main():
    # ''' This Section uses multi-threading to send config to ten TP endpoint at a time'''
    # pool = ThreadPool(10)
    # results = pool.map(do_upload, theDestinationDevices)
    # pool.close()
    # pool.join()
    # return results
    for someDevice in theDestinationDevices:
        do_upload(someDevice)

main()











