"""Written for creating a Dormat Users list for OCI IAM"""

import json
import requests
import base64
import urllib3
import xlsxwriter
import os
import datetime
urllib3.disable_warnings()
requests.packages.urllib3.util.ssl_ = 'ALL:@SECLEVEL=1'

class IAM():

    #import config files
    def __init__(self):
        config = json.load(open('config.json'))
        global idcsURL
        global clientID
        global clientSecret

        idcsURL = config["iamurl"]
        clientID = config["client_id"]
        clientSecret = config["client_secret"]

    #encode client & secret
    def get_encoded(self,clid, clsecret):    #6.
        encoded = clid + ":" + clsecret
        baseencoded = base64.urlsafe_b64encode(encoded.encode('UTF-8')).decode('ascii')
        return baseencoded

    #get access token
    def get_access_token(self,url, header):    #8.
        para = "grant_type=client_credentials&scope=urn:opc:idm:__myscopes__"
        response = requests.post(url, headers=header, data=para, verify=False)
        jsonresp = json.loads(response.content)
        access_token = jsonresp.get('access_token')
        return access_token

    #print access token
    def printaccesstoken(self):  #4.
        obj = IAM()
        encodedtoken = obj.get_encoded(clientID, clientSecret)     #5.
        extra = "/oauth2/v1/token"
        headers = {'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8',
                   'Authorization': 'Basic %s' % encodedtoken, 'Accept': '*/*'}
        accesstoken = obj.get_access_token(idcsURL + extra, headers)     #7.
        return accesstoken

    def get_successfullogindate(self):
        extra = "/admin/v1/Users"
        obj = IAM()
        accesstoken = obj.printaccesstoken()
        headers = {'Authorization': 'Bearer ' + accesstoken}
        resp = requests.get(idcsURL+extra, headers=headers, verify=False)
        jsonresp = json.loads(resp.content)
        totalCount = jsonresp.get("totalResults")
        print("Total number of users: " + str(totalCount))
        
        wb=xlsxwriter.Workbook('Dormat_Users.xlsx')
        sheet= wb.add_worksheet()
        sheet.write('A1', 'DormantUsers_UserName')
        sheet.write('B1', 'DormantUsers_LastSuccessfulLogin')
        sheet.write('C1', 'DormantUsers_FullName')
        sheet.write('D1', 'DormantUsers_Created_On')
        sheet.write('E1', 'Users_Status')
        sheet.write('F1', 'Remark')
       

        current_date = datetime.datetime.now().date()
        startIndex = 1
        count = 50
        row = 1
        column = 0
        loop = int(totalCount / count)
        extra2 = "/admin/v1/AppRoles"

        for i in range(loop + 1):
            param = {'attributes': "displayName,username,meta,active,urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User:lastSuccessfulLoginDate",
                     'startIndex': startIndex, 'count': count}
            resp = requests.get(idcsURL+extra, headers=headers, verify=False, params=param)
            startIndex += count
            jsonresp = json.loads(resp.content)

            tempjsn = jsonresp.get("Resources")

            for x in tempjsn:
                trimjsn = {}
                username = trimjsn['userName'] = x.get("userName")
                userid = x.get("id")
                userStatus = x.get("active")
                displayName = x.get("displayName")
                userState = x.get("urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User",{})
                createdOn=x['meta']['created']
                lastLoginDate = userState.get('lastSuccessfulLoginDate')
                if userStatus is False:
                    sheet.write(row, column, username)
                    sheet.write(row, column+1, lastLoginDate)
                    sheet.write(row, column+2, displayName)
                    sheet.write(row, column+3, createdOn)
                    sheet.write(row, column+4, userStatus)
                    sheet.write(row, column+5, "User is InActive. Will be Deleted")
                    row += 1
                    
                    #print(username + " is in inactive state. Will be deleted.")
                    with open('username.txt', 'a') as c:
                        c.write(username)
                        c.write('\n')
                    with open('userID.txt', 'a') as d:
                        d.write(userid)
                        d.write('\n')
                
                    continue
                if userState is None:
                    
                    sheet.write(row, column, username)
                    sheet.write(row, column+1, "")
                    sheet.write(row, column+2, displayName)
                    sheet.write(row, column+3, createdOn)
                    sheet.write(row, column+4, userStatus)
                    sheet.write(row, column+5, "User has never accessed. Kindly make a decision to delete it manually on console")
                    row += 1
                    
                    #print(username + " has never accessed. Kindly make a decision to delete it manually on console")
                    # with open('username.txt', 'a') as c:
                    #     c.write(username)
                    #     c.write('\n')
                    # with open('userID.txt', 'a') as d:
                    #     d.write(userid)
                    #     d.write('\n')

                    #  ** Uncomment the lines commented above to delete the users who have never accessed **

                    continue
                
                if lastLoginDate:
                    target_dates = datetime.datetime.strptime(lastLoginDate, '%Y-%m-%dT%H:%M:%S.%fZ').date()
                    num_days = (current_date - target_dates).days
                else:
                    target_dates = None
                    num_days = None
                
                if num_days is not None and num_days > 90:    
                    #print("entered if condition for more than 90 days ")
                    params2 = {"filter": f'members[value eq "{userid}"] and app.value eq "IDCSAppId" and (displayName eq "Identity Domain Administrator" or displayName eq "Security Administrator" or displayName eq "Application Administrator" or displayName eq "User Administrator" or displayName eq "User Manager" or displayName eq "Help Desk Administrator" or displayName eq "Audit Administrator")'}
                    resp2 = requests.get(idcsURL+extra2, headers=headers, verify=False, params=params2)
                    jsonresp3 = json.loads(resp2.content)
                    totalCount1 = jsonresp3.get("totalResults")
                    if totalCount1 < 1:
                        
                        sheet.write(row, column, username)
                        sheet.write(row, column+1, lastLoginDate)
                        sheet.write(row, column+2, displayName)
                        sheet.write(row, column+3, createdOn)
                        sheet.write(row, column+4, userStatus)
                        sheet.write(row, column+5, "User has not Logged in for 90 days and has NO admin role. Will be deleted")
                        row += 1
                        
                        #print(username + " has not Logged in for 90 days and has NO admin role. Will be deleted.")
                        with open('username.txt', 'a') as c:
                            c.write(username)
                            c.write('\n')
                        with open('userID.txt', 'a') as d:
                            d.write(userid)
                            d.write('\n')

        wb.close()
        try:
            with open('userID.txt','r') as g:
                lines = len(g.readlines())
        except FileNotFoundError:
            print("No users found to be Deleted by the Script")
            exit()

        x=0
        while x<lines:
            with open('userID.txt','r') as e:
                content = e.readlines()
                id=content[x]
                id=id.rstrip(id[-1])
            paradelete=json.dumps(
                    {
                    "method": "DELETE",
                    "path": "/Users/"+id+"?forceDelete=true"
                    }
                    )
            with open('username.json', 'a') as j:
                    j.writelines(paradelete)
                    j.write(',')

            x=x+1
        with open('username.json', 'r') as file:
            data = file.read()
        with open('formatedData.json', 'a') as file1:
            file1.writelines('{ \n')
            file1.writelines('"schemas": ["urn:ietf:params:scim:api:messages:2.0:BulkRequest"],\n')
            file1.writelines('"Operations": [\n')
            data2=data.rstrip(data[-1])
            file1.write(data2)
            file1.write(']}')

        headers = {'Authorization': 'Bearer ' + accesstoken}
        bulkdata = json.load(open('formatedData.json'))
        param = {'forceDelete': True}
        payload = json.dumps(bulkdata)
        headers2 = {'Content-Type': 'application/json','Authorization': 'Bearer ' + accesstoken, 'Accept': '*/*'}
        extra2="/admin/v1/Bulk"
        #respdelete = requests.request("POST",idcsURL + extra2, headers=headers2, verify=False, params=param,data=payload)
        #print("The identified users has been deleted ")

        # try:
        #     os.remove('userID.txt')
        #     os.remove('username.txt')
        #     os.remove('formatedData.json')
        #     os.remove('username.json')
        # except FileNotFoundError:
        #     exit()

obj = IAM()   #create an object
obj.get_successfullogindate()
                  
