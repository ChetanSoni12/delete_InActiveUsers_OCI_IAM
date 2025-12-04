"""Export all OCI IAM Users to Excel"""

import json
import requests
import base64
import urllib3
import xlsxwriter
import datetime

urllib3.disable_warnings()

class IAM:

    def __init__(self):
        config = json.load(open('config.json'))
        self.idcsURL = config["iamurl"]
        self.clientID = config["client_id"]
        self.clientSecret = config["client_secret"]

    def get_encoded(self):
        raw = f"{self.clientID}:{self.clientSecret}"
        return base64.urlsafe_b64encode(raw.encode()).decode()

    def get_access_token(self):
        url = self.idcsURL + "/oauth2/v1/token"
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Authorization": f"Basic {self.get_encoded()}"
        }
        data = "grant_type=client_credentials&scope=urn:opc:idm:__myscopes__"

        resp = requests.post(url, headers=headers, data=data, verify=False)
        return resp.json().get("access_token")

    def export_all_users(self):

        access_token = self.get_access_token()
        headers = {"Authorization": "Bearer " + access_token}

        extra = "/admin/v1/Users"
        wb = xlsxwriter.Workbook("All_Users.xlsx")
        sheet = wb.add_worksheet()

        # Columns
        columns = [
            "UserName",
            "DisplayName",
            "Email",
            "ID",
            "Active",
            "IsFederatedUser",
            "CreatedOn",
            "LastModified",
            "LastLoginDate"
        ]

        for col, c in enumerate(columns):
            sheet.write(0, col, c)

        row = 1
        start_index = 1
        count = 50

        print("Collecting user count ...")
        first_call = requests.get(self.idcsURL + extra, headers=headers, verify=False)
        total_users = first_call.json().get("totalResults", 0)
        print(f"Total users: {total_users}")

        loops = int(total_users / count) + 1

        for _ in range(loops):
            params = {
                "attributes":
                "userName,displayName,emails,active,meta,"
                "urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User:lastSuccessfulLoginDate,"
                "urn:ietf:params:scim:schemas:oracle:idcs:extension:user:User:isFederatedUser",
                "startIndex": start_index,
                "count": count
            }

            resp = requests.get(self.idcsURL + extra, headers=headers, params=params, verify=False)
            start_index += count
            items = resp.json().get("Resources", [])

            for u in items:
                emails = u.get("emails", [])
                email = emails[0]["value"] if emails else ""

                # User State (Last Login)
                user_state = u.get(
                    "urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User",
                    {}
                )
                last_login = user_state.get("lastSuccessfulLoginDate", "")

                # Federated Flag
                fed_ext = u.get(
                    "urn:ietf:params:scim:schemas:oracle:idcs:extension:user:User",
                    {}
                )
                is_fed_user = fed_ext.get("isFederatedUser", "")

                data_row = [
                    u.get("userName", ""),
                    u.get("displayName", ""),
                    email,
                    u.get("id", ""),
                    u.get("active", ""),
                    is_fed_user,
                    u.get("meta", {}).get("created", ""),
                    u.get("meta", {}).get("lastModified", ""),
                    last_login
                ]

                for col, val in enumerate(data_row):
                    sheet.write(row, col, val)

                row += 1

        wb.close()
        print("Completed: All_Users.xlsx created.")

# Run
obj = IAM()
obj.export_all_users()
