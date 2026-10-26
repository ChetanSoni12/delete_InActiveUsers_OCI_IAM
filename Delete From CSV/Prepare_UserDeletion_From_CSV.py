"""Prepare OCI IAM user deletions from a CSV username list.

Expected CSV format:

    username
    user1@example.com
    user2@example.com

The script is dry-run only by default. It validates usernames against OCI IAM,
creates an Excel review report, and writes bulk payload previews in batches of
no more than 500 operations. The live delete block is commented out.
"""

import argparse
import base64
import csv
import datetime
import json
from pathlib import Path

import requests
import urllib3
import xlsxwriter

urllib3.disable_warnings()
requests.packages.urllib3.util.ssl_ = "ALL:@SECLEVEL=1"


class CSVUserDeletePreparer:
    BULK_DELETE_LIMIT = 500
    USER_PAGE_SIZE = 200
    REQUEST_TIMEOUT = (10, 600)
    PREVIEW_DIR = Path("bulk_delete_payload_preview_csv_users")
    REVIEW_FILE = Path("CSV_User_Deletion_Review.xlsx")
    LOG_FILE = Path("prepare_csv_user_deletion.log")
    EXECUTION_LOG_FILE = Path("execute_csv_user_deletion.log")
    USERNAME_HEADERS = {"username", "user_name", "user name", "email"}

    def __init__(self, csv_path):
        self.csv_path = Path(csv_path)

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

    def _encode_credentials(self):
        credentials = self.client_id + ":" + self.client_secret
        return base64.urlsafe_b64encode(credentials.encode("utf-8")).decode("ascii")

    def _build_access_token(self):
        headers = {
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "Authorization": "Basic " + self._encode_credentials(),
            "Accept": "*/*",
        }
        response = self.session.post(
            self.idcs_url + "/oauth2/v1/token",
            headers=headers,
            data="grant_type=client_credentials&scope=urn:opc:idm:__myscopes__",
            timeout=self.REQUEST_TIMEOUT,
        )
        response.raise_for_status()
        access_token = response.json().get("access_token")
        if not access_token:
            raise ValueError("Unable to retrieve access token from IAM response.")
        return access_token

    def _read_usernames(self):
        if not self.csv_path.is_file():
            raise FileNotFoundError("CSV file not found: " + str(self.csv_path))

        with self.csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            reader = csv.DictReader(csv_file)
            if not reader.fieldnames:
                raise ValueError("CSV is empty or does not contain a header row.")

            username_column = None
            for field_name in reader.fieldnames:
                if field_name and field_name.strip().casefold() in self.USERNAME_HEADERS:
                    username_column = field_name
                    break

            if username_column is None:
                raise ValueError(
                    "CSV must contain a username column. Supported headers: username, user_name, user name, email."
                )

            usernames = []
            duplicate_usernames = []
            seen_keys = set()

            for row_number, row in enumerate(reader, start=2):
                username = (row.get(username_column) or "").strip()
                if not username:
                    continue

                username_key = username.casefold()
                if username_key in seen_keys:
                    duplicate_usernames.append(username)
                    continue

                seen_keys.add(username_key)
                usernames.append(username)

        if not usernames:
            raise ValueError("CSV does not contain any non-empty usernames.")

        return usernames, duplicate_usernames

    def _find_users(self, requested_usernames):
        requested_by_key = {username.casefold(): username for username in requested_usernames}
        matched_by_key = {}
        ambiguous_keys = set()
        start_index = 1
        total_count = None
        processed_users = 0

        print("Looking up {count} unique CSV username(s) in OCI IAM...".format(count=len(requested_usernames)))

        while total_count is None or start_index <= total_count:
            response = self.session.get(
                self.idcs_url + "/admin/v1/Users",
                headers=self.auth_headers,
                params={
                    "attributes": "id,userName,displayName,active",
                    "startIndex": start_index,
                    "count": self.USER_PAGE_SIZE,
                },
                timeout=self.REQUEST_TIMEOUT,
            )
            response.raise_for_status()
            response_json = response.json()
            total_count = response_json.get("totalResults", 0)
            resources = response_json.get("Resources", [])
            if not resources:
                break

            for user in resources:
                processed_users += 1
                username = (user.get("userName") or "").strip()
                username_key = username.casefold()
                if username_key not in requested_by_key:
                    continue

                previous_match = matched_by_key.get(username_key)
                if previous_match and previous_match.get("id") != user.get("id"):
                    ambiguous_keys.add(username_key)
                    matched_by_key.pop(username_key, None)
                    continue

                if username_key not in ambiguous_keys:
                    matched_by_key[username_key] = {
                        "requested_username": requested_by_key[username_key],
                        "username": username,
                        "id": user.get("id"),
                        "display_name": user.get("displayName"),
                        "active": user.get("active"),
                    }

            resolved_count = len(matched_by_key) + len(ambiguous_keys)
            print(
                "Scanned {processed} of {total} users. Resolved {resolved} of {requested} CSV username(s).".format(
                    processed=min(processed_users, total_count),
                    total=total_count,
                    resolved=resolved_count,
                    requested=len(requested_usernames),
                )
            )
            if resolved_count == len(requested_usernames):
                break

            start_index += self.USER_PAGE_SIZE

        matched_users = []
        missing_usernames = []
        ambiguous_usernames = []

        for requested_username in requested_usernames:
            username_key = requested_username.casefold()
            if username_key in ambiguous_keys:
                ambiguous_usernames.append(requested_username)
            elif username_key in matched_by_key:
                matched_users.append(matched_by_key[username_key])
            else:
                missing_usernames.append(requested_username)

        return matched_users, missing_usernames, ambiguous_usernames, total_count

    def _write_review_workbook(self, requested_usernames, matched_users, missing_usernames, ambiguous_usernames):
        matched_by_key = {user["requested_username"].casefold(): user for user in matched_users}
        missing_keys = {username.casefold() for username in missing_usernames}
        ambiguous_keys = {username.casefold() for username in ambiguous_usernames}
        workbook = xlsxwriter.Workbook(str(self.REVIEW_FILE))
        sheet = workbook.add_worksheet("CSV User Review")
        headers = [
            "CSV_Username",
            "OCI_Username",
            "OCI_User_ID",
            "Display_Name",
            "Active",
            "Validation_Status",
            "Planned_Action",
        ]

        for column, header in enumerate(headers):
            sheet.write_string(0, column, header)

        for row_number, requested_username in enumerate(requested_usernames, start=1):
            username_key = requested_username.casefold()
            matched_user = matched_by_key.get(username_key)
            values = [requested_username, "", "", "", "", "", ""]

            if matched_user:
                values = [
                    requested_username,
                    matched_user["username"],
                    matched_user["id"] or "",
                    matched_user["display_name"] or "",
                    str(matched_user["active"]),
                    "MATCHED",
                    "Will be included in delete batch preview",
                ]
            elif username_key in ambiguous_keys:
                values[5] = "AMBIGUOUS"
                values[6] = "Skipped for safety"
            elif username_key in missing_keys:
                values[5] = "NOT_FOUND"
                values[6] = "Skipped"

            for column, value in enumerate(values):
                sheet.write_string(row_number, column, str(value))

        sheet.freeze_panes(1, 0)
        sheet.autofilter(0, 0, len(requested_usernames), len(headers) - 1)
        sheet.set_column(0, 1, 35)
        sheet.set_column(2, 2, 40)
        sheet.set_column(3, 3, 30)
        sheet.set_column(4, 6, 28)
        workbook.close()

    def _build_delete_operations(self, matched_users):
        return [
            {
                "method": "DELETE",
                "path": "/Users/" + user["id"] + "?forceDelete=true",
                "bulkId": user["id"],
            }
            for user in matched_users
            if user.get("id")
        ]

    def _chunk_operations(self, operations):
        for index in range(0, len(operations), self.BULK_DELETE_LIMIT):
            yield operations[index:index + self.BULK_DELETE_LIMIT]

    def _write_preview_batches(self, operations):
        self.PREVIEW_DIR.mkdir(exist_ok=True)
        for existing_file in self.PREVIEW_DIR.glob("batch_*.json"):
            existing_file.unlink()

        batch_count = 0
        for batch_number, operation_batch in enumerate(self._chunk_operations(operations), start=1):
            payload = {
                "schemas": ["urn:ietf:params:scim:api:messages:2.0:BulkRequest"],
                "Operations": operation_batch,
            }
            batch_file = self.PREVIEW_DIR / "batch_{number:03d}.json".format(number=batch_number)
            batch_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            batch_count = batch_number

        return batch_count

    def _write_preparation_log(self, requested_count, duplicate_usernames, matched_users, missing_usernames, ambiguous_usernames, batch_count, domain_user_count):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.LOG_FILE.open("a", encoding="utf-8") as log_file:
            log_file.write("Execution Timestamp: " + timestamp + "\n")
            log_file.write("CSV File: " + str(self.csv_path) + "\n")
            log_file.write("Domain Users Reported: " + str(domain_user_count) + "\n")
            log_file.write("Unique CSV Usernames: " + str(requested_count) + "\n")
            log_file.write("Duplicate CSV Rows Skipped: " + str(len(duplicate_usernames)) + "\n")
            log_file.write("Matched Users Prepared: " + str(len(matched_users)) + "\n")
            log_file.write("Missing Users: " + str(len(missing_usernames)) + "\n")
            log_file.write("Ambiguous Users Skipped: " + str(len(ambiguous_usernames)) + "\n")
            log_file.write("Batches Prepared: " + str(batch_count) + "\n")
            log_file.write("Dry Run: Yes\n")
            for username in missing_usernames:
                log_file.write("NOT_FOUND: " + username + "\n")
            for username in ambiguous_usernames:
                log_file.write("AMBIGUOUS_SKIPPED: " + username + "\n")
            log_file.write("-" * 80 + "\n")

    def _get_operation_status(self, operation_result):
        status = operation_result.get("status", "")
        if isinstance(status, dict):
            status = status.get("code", "")
        return str(status)

    def _get_operation_detail(self, operation_result):
        detail = operation_result.get("detail", "")
        response_data = operation_result.get("response")
        if not detail and isinstance(response_data, dict):
            detail = response_data.get("detail", "")
        if not detail and response_data:
            detail = response_data
        return detail

    def _report_bulk_response(self, batch_number, response, requested_operations):
        try:
            response_json = response.json()
        except ValueError:
            response_json = {}

        operation_results = response_json.get("Operations", [])
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        succeeded = 0
        failed = 0

        with self.EXECUTION_LOG_FILE.open("a", encoding="utf-8") as log_file:
            log_file.write("Execution Timestamp: " + timestamp + "\n")
            log_file.write("Batch: " + str(batch_number) + "\n")
            log_file.write("Bulk HTTP Status: " + str(response.status_code) + "\n")

            if not operation_results:
                detail = response.text[:4000]
                log_file.write("No per-operation results returned. Response: " + detail + "\n")
                log_file.write("-" * 80 + "\n")
                print("OCI returned no per-user operation statuses. Check " + str(self.EXECUTION_LOG_FILE))
                return False

            for index, operation_result in enumerate(operation_results):
                status = self._get_operation_status(operation_result)
                detail = self._get_operation_detail(operation_result)
                requested_operation = requested_operations[index] if index < len(requested_operations) else {}
                user_id = operation_result.get("bulkId") or requested_operation.get("bulkId", "UNKNOWN")
                path = operation_result.get("path") or requested_operation.get("path", "")
                is_success = status.startswith("2")

                if is_success:
                    succeeded += 1
                else:
                    failed += 1

                result_line = (
                    "user_id={user_id}, status={status}, path={path}, detail={detail}"
                ).format(
                    user_id=user_id,
                    status=status or "UNKNOWN",
                    path=path,
                    detail=json.dumps(detail),
                )
                log_file.write(result_line + "\n")
                print(result_line)

            log_file.write(
                "Succeeded: {succeeded}, Failed: {failed}\n".format(
                    succeeded=succeeded,
                    failed=failed,
                )
            )
            log_file.write("-" * 80 + "\n")

        print(
            "Batch {batch} result: {succeeded} succeeded, {failed} failed.".format(
                batch=batch_number,
                succeeded=succeeded,
                failed=failed,
            )
        )
        return failed == 0 and succeeded == len(requested_operations)

    def prepare(self):
        requested_usernames, duplicate_usernames = self._read_usernames()
        matched_users, missing_usernames, ambiguous_usernames, domain_user_count = self._find_users(
            requested_usernames
        )
        self._write_review_workbook(
            requested_usernames,
            matched_users,
            missing_usernames,
            ambiguous_usernames,
        )
        operations = self._build_delete_operations(matched_users)
        batch_count = self._write_preview_batches(operations)
        self._write_preparation_log(
            len(requested_usernames),
            duplicate_usernames,
            matched_users,
            missing_usernames,
            ambiguous_usernames,
            batch_count,
            domain_user_count,
        )

        print("CSV usernames requested: " + str(len(requested_usernames)))
        print("Matched users prepared: " + str(len(matched_users)))
        print("Missing usernames: " + str(len(missing_usernames)))
        print("Duplicate CSV rows skipped: " + str(len(duplicate_usernames)))
        print("Bulk preview batches prepared: " + str(batch_count))
        print("Preparation complete. No delete request has been submitted by the preparation stage.")

        # To enable deletion, remove the leading '# ' from every line between
        # LIVE DELETE BLOCK START and LIVE DELETE BLOCK END. Review the Excel
        # report and JSON batches before doing so.
        # LIVE DELETE BLOCK START
        # for batch_number, operation_batch in enumerate(self._chunk_operations(operations), start=1):
        #     payload = {
        #         "schemas": ["urn:ietf:params:scim:api:messages:2.0:BulkRequest"],
        #         "Operations": operation_batch,
        #     }
        #     print(
        #         "Submitting batch {batch} with {count} delete operation(s)...".format(
        #             batch=batch_number,
        #             count=len(operation_batch),
        #         )
        #     )
        #     response = self.session.post(
        #         self.idcs_url + "/admin/v1/Bulk",
        #         headers=self.bulk_headers,
        #         params={"forceDelete": True},
        #         data=json.dumps(payload),
        #         timeout=self.REQUEST_TIMEOUT,
        #     )
        #     response.raise_for_status()
        #     batch_succeeded = self._report_bulk_response(
        #         batch_number,
        #         response,
        #         operation_batch,
        #     )
        #     if not batch_succeeded:
        #         print("Stopping because one or more delete operations failed or had an unknown result.")
        #         break
        # LIVE DELETE BLOCK END


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Validate OCI IAM usernames from CSV and prepare delete batches."
    )
    parser.add_argument(
        "csv_file",
        nargs="?",
        default="usernames.csv",
        help="CSV input path. Default: usernames.csv",
    )
    return parser.parse_args()


arguments = parse_arguments()
preparer = CSVUserDeletePreparer(arguments.csv_file)
preparer.prepare()