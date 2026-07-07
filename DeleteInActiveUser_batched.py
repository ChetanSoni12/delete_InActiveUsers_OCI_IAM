"""Written for creating a Dormant Users list for OCI IAM and deleting in batches."""

import base64
import datetime
import json
from pathlib import Path

import requests
import urllib3
import xlsxwriter

urllib3.disable_warnings()
requests.packages.urllib3.util.ssl_ = "ALL:@SECLEVEL=1"


class IAM:
    BULK_DELETE_LIMIT = 500
    USER_PAGE_SIZE = 50
    INACTIVE_DAYS_THRESHOLD = 90
    ADMIN_ROLE_FILTER = (
        'members[value eq "{user_id}"] and app.value eq "IDCSAppId" and '
        '(displayName eq "Identity Domain Administrator" or '
        'displayName eq "Security Administrator" or '
        'displayName eq "Application Administrator" or '
        'displayName eq "User Administrator" or '
        'displayName eq "User Manager" or '
        'displayName eq "Help Desk Administrator" or '
        'displayName eq "Audit Administrator")'
    )

    def __init__(self):
        with open("config.json", "r", encoding="utf-8") as config_file:
            config = json.load(config_file)

        self.idcs_url = config["iamurl"]
        self.client_id = config["client_id"]
        self.client_secret = config["client_secret"]
        self.session = requests.Session()
        self.session.verify = False
        self.access_token = self._build_access_token()
        self.auth_headers = {"Authorization": "Bearer " + self.access_token}
        self.bulk_headers = {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + self.access_token,
            "Accept": "*/*",
        }

    def get_encoded(self, client_id, client_secret):
        encoded = client_id + ":" + client_secret
        return base64.urlsafe_b64encode(encoded.encode("utf-8")).decode("ascii")

    def get_access_token(self, url, header):
        params = "grant_type=client_credentials&scope=urn:opc:idm:__myscopes__"
        response = self.session.post(url, headers=header, data=params)
        response.raise_for_status()
        json_response = response.json()
        access_token = json_response.get("access_token")
        if not access_token:
            raise ValueError("Unable to retrieve access token from IAM response.")
        return access_token

    def _build_access_token(self):
        encoded_token = self.get_encoded(self.client_id, self.client_secret)
        headers = {
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "Authorization": "Basic %s" % encoded_token,
            "Accept": "*/*",
        }
        return self.get_access_token(self.idcs_url + "/oauth2/v1/token", headers)

    def _build_workbook(self):
        workbook = xlsxwriter.Workbook("Dormat_Users.xlsx")
        sheet = workbook.add_worksheet()
        headers = [
            "DormantUsers_UserName",
            "DormantUsers_LastSuccessfulLogin",
            "DormantUsers_FullName",
            "DormantUsers_Created_On",
            "Users_Status",
            "Remark",
        ]

        for column, header in enumerate(headers):
            sheet.write(0, column, header)

        return workbook, sheet

    def _append_user_row(self, sheet, row_number, username, last_login_date, display_name, created_on, user_status, remark):
        sheet.write(row_number, 0, username)
        sheet.write(row_number, 1, last_login_date or "")
        sheet.write(row_number, 2, display_name)
        sheet.write(row_number, 3, created_on)
        sheet.write(row_number, 4, user_status)
        sheet.write(row_number, 5, remark)

    def _is_admin_user(self, user_id):
        response = self.session.get(
            self.idcs_url + "/admin/v1/AppRoles",
            headers=self.auth_headers,
            params={"filter": self.ADMIN_ROLE_FILTER.format(user_id=user_id)},
        )
        response.raise_for_status()
        json_response = response.json()
        return json_response.get("totalResults", 0) >= 1

    def _build_delete_operations(self, users_to_delete):
        operations = []
        for user in users_to_delete:
            operations.append(
                {
                    "method": "DELETE",
                    "path": "/Users/" + user["id"] + "?forceDelete=true",
                    "bulkId": user["id"],
                }
            )
        return operations

    def _chunk_operations(self, operations, chunk_size):
        for index in range(0, len(operations), chunk_size):
            yield operations[index:index + chunk_size]

    def _extract_bulk_failures(self, response_json, user_lookup):
        failed_users = []

        for operation_result in response_json.get("Operations", []):
            status_value = str(operation_result.get("status", ""))
            if status_value.startswith("2"):
                continue

            bulk_id = operation_result.get("bulkId")
            matched_user = user_lookup.get(bulk_id, {"id": bulk_id, "username": "UNKNOWN"})
            failed_users.append(
                {
                    "user_id": matched_user["id"],
                    "username": matched_user["username"],
                    "status": status_value or "UNKNOWN",
                    "detail": operation_result.get("response", ""),
                }
            )

        return failed_users

    def _write_failed_user_log(self, failed_users, batches_sent):
        log_path = Path("delete_inactive_users_execution.log")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with log_path.open("a", encoding="utf-8") as log_file:
            log_file.write("Execution Timestamp: " + timestamp + "\n")
            log_file.write("Batches Submitted: " + str(batches_sent) + "\n")

            if failed_users:
                log_file.write("Failed Users:\n")
                for failed_user in failed_users:
                    log_file.write(
                        "username={username}, user_id={user_id}, status={status}, detail={detail}\n".format(
                            username=failed_user["username"],
                            user_id=failed_user["user_id"],
                            status=failed_user["status"],
                            detail=json.dumps(failed_user["detail"]),
                        )
                    )
            else:
                log_file.write("Failed Users: None\n")

            log_file.write("-" * 80 + "\n")

    def _write_bulk_payload_preview(self, operations):
        payload_path = Path("bulk_delete_payload_preview.json")
        payload = {
            "schemas": ["urn:ietf:params:scim:api:messages:2.0:BulkRequest"],
            "Operations": operations,
        }
        payload_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    def _delete_users_in_batches(self, users_to_delete):
        if not users_to_delete:
            print("No users found to be Deleted by the Script")
            return

        operations = self._build_delete_operations(users_to_delete)
        self._write_bulk_payload_preview(operations)
        batch_count = len(list(self._chunk_operations(operations, self.BULK_DELETE_LIMIT)))
        self._write_failed_user_log([], 0)
        print(
            "Prepared {users} delete operations in {batches} batch(es) of up to {limit}. "
            "Actual bulk deletion is intentionally disabled.".format(
                users=len(users_to_delete),
                batches=batch_count,
                limit=self.BULK_DELETE_LIMIT,
            )
        )

        # Uncomment the block below to enable actual deletion in batches of up to 500 users.
        # user_lookup = {user["id"]: user for user in users_to_delete}
        # failed_users = []
        # batches_sent = 0
        #
        # for operation_batch in self._chunk_operations(operations, self.BULK_DELETE_LIMIT):
        #     payload = {
        #         "schemas": ["urn:ietf:params:scim:api:messages:2.0:BulkRequest"],
        #         "Operations": operation_batch,
        #     }
        #     response = self.session.post(
        #         self.idcs_url + "/admin/v1/Bulk",
        #         headers=self.bulk_headers,
        #         params={"forceDelete": True},
        #         data=json.dumps(payload),
        #     )
        #     batches_sent += 1
        #
        #     if response.ok:
        #         response_json = response.json()
        #         failed_users.extend(self._extract_bulk_failures(response_json, user_lookup))
        #     else:
        #         for operation in operation_batch:
        #             matched_user = user_lookup.get(
        #                 operation["bulkId"],
        #                 {"id": operation["bulkId"], "username": "UNKNOWN"},
        #             )
        #             failed_users.append(
        #                 {
        #                     "user_id": matched_user["id"],
        #                     "username": matched_user["username"],
        #                     "status": str(response.status_code),
        #                     "detail": response.text,
        #                 }
        #             )
        #
        # self._write_failed_user_log(failed_users, batches_sent)
        # print(
        #     "Deletion batches submitted: {batches}. Failed deletions logged: {failed}".format(
        #         batches=batches_sent,
        #         failed=len(failed_users),
        #     )
        # )

    def get_successfullogindate(self):
        response = self.session.get(self.idcs_url + "/admin/v1/Users", headers=self.auth_headers)
        response.raise_for_status()
        json_response = response.json()
        total_count = json_response.get("totalResults", 0)
        print("Total number of users: " + str(total_count))

        workbook, sheet = self._build_workbook()
        current_date = datetime.datetime.now().date()
        start_index = 1
        row = 1
        users_to_delete = []

        while start_index <= total_count:
            params = {
                "attributes": (
                    "displayName,username,meta,active,"
                    "urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User:lastSuccessfulLoginDate"
                ),
                "startIndex": start_index,
                "count": self.USER_PAGE_SIZE,
            }
            response = self.session.get(
                self.idcs_url + "/admin/v1/Users",
                headers=self.auth_headers,
                params=params,
            )
            response.raise_for_status()
            json_response = response.json()
            resources = json_response.get("Resources", [])
            if not resources:
                break

            for user in resources:
                username = user.get("userName")
                user_id = user.get("id")
                user_status = user.get("active")
                display_name = user.get("displayName")
                user_state = user.get(
                    "urn:ietf:params:scim:schemas:oracle:idcs:extension:userState:User",
                    {},
                )
                created_on = user["meta"]["created"]

                if user_status is False:
                    last_login_date = user_state.get("lastSuccessfulLoginDate") if user_state else ""
                    self._append_user_row(
                        sheet,
                        row,
                        username,
                        last_login_date,
                        display_name,
                        created_on,
                        user_status,
                        "User is InActive. Will be Deleted",
                    )
                    users_to_delete.append({"id": user_id, "username": username})
                    row += 1
                    continue

                if user_state is None:
                    self._append_user_row(
                        sheet,
                        row,
                        username,
                        "",
                        display_name,
                        created_on,
                        user_status,
                        "User has never accessed. Kindly make a decision to delete it manually on console",
                    )
                    row += 1
                    continue

                last_login_date = user_state.get("lastSuccessfulLoginDate")
                if last_login_date:
                    target_date = datetime.datetime.strptime(
                        last_login_date, "%Y-%m-%dT%H:%M:%S.%fZ"
                    ).date()
                    days_inactive = (current_date - target_date).days
                else:
                    days_inactive = None

                if days_inactive is not None and days_inactive > self.INACTIVE_DAYS_THRESHOLD:
                    if not self._is_admin_user(user_id):
                        self._append_user_row(
                            sheet,
                            row,
                            username,
                            last_login_date,
                            display_name,
                            created_on,
                            user_status,
                            "User has not Logged in for 90 days and has NO admin role. Will be deleted",
                        )
                        users_to_delete.append({"id": user_id, "username": username})
                        row += 1

            start_index += self.USER_PAGE_SIZE

        workbook.close()
        self._delete_users_in_batches(users_to_delete)


obj = IAM()
obj.get_successfullogindate()
