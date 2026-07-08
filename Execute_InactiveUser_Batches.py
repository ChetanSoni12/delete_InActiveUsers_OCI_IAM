"""Execute a limited number of prepared inactive-user delete batches."""

import base64
import datetime
import json
from pathlib import Path

import requests
import urllib3

urllib3.disable_warnings()
requests.packages.urllib3.util.ssl_ = "ALL:@SECLEVEL=1"


class BatchDeleteExecutor:
    REQUEST_TIMEOUT = (10, 120)
    PREVIEW_DIR = Path("bulk_delete_payload_preview_inactive_only")
    LOG_FILE = Path("execute_inactive_only_batches.log")

    def __init__(self):
        with open("config.json", "r", encoding="utf-8") as config_file:
            config = json.load(config_file)

        self.idcs_url = config["iamurl"]
        self.client_id = config["client_id"]
        self.client_secret = config["client_secret"]
        self.session = requests.Session()
        self.session.verify = False
        self.access_token = self._build_access_token()
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
        response = self.session.post(url, headers=header, data=params, timeout=self.REQUEST_TIMEOUT)
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

    def _get_batch_files(self):
        if not self.PREVIEW_DIR.exists():
            raise FileNotFoundError(
                "Preview directory not found: {path}".format(path=self.PREVIEW_DIR)
            )

        batch_files = sorted(self.PREVIEW_DIR.glob("batch_*.json"))
        if not batch_files:
            raise FileNotFoundError(
                "No batch preview files found in: {path}".format(path=self.PREVIEW_DIR)
            )

        return batch_files

    def _prompt_batch_start(self, total_batches):
        while True:
            user_input = input(
                "Enter the starting batch number to execute [1-{total}]: ".format(
                    total=total_batches,
                )
            ).strip()
            if user_input.isdigit():
                start_batch = int(user_input)
                if 1 <= start_batch <= total_batches:
                    return start_batch
            print("Please enter a whole number between 1 and {total}.".format(total=total_batches))

    def _prompt_batch_count(self, max_batches):
        while True:
            user_input = input(
                "Enter how many batches to execute from that point [1-{total}]: ".format(
                    total=max_batches,
                )
            ).strip()
            if user_input.isdigit():
                batch_count = int(user_input)
                if 1 <= batch_count <= max_batches:
                    return batch_count
            print("Please enter a whole number between 1 and {total}.".format(total=max_batches))

    def _prompt_confirmation(self, selected_files, selected_operations):
        print(
            "You are about to execute {batches} batch(es) covering {users} delete operation(s).".format(
                batches=len(selected_files),
                users=selected_operations,
            )
        )
        print(
            "Batch range: {start} to {end}".format(
                start=selected_files[0].name,
                end=selected_files[-1].name,
            )
        )
        confirmation = input('Type "DELETE" to continue: ').strip()
        return confirmation == "DELETE"

    def _write_execution_log(self, selected_files, batch_results):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.LOG_FILE.open("a", encoding="utf-8") as log_file:
            log_file.write("Execution Timestamp: " + timestamp + "\n")
            log_file.write("Batches Requested: " + str(len(selected_files)) + "\n")
            log_file.write(
                "Batch Files: " + ", ".join(batch_file.name for batch_file in selected_files) + "\n"
            )
            for batch_result in batch_results:
                log_file.write(
                    "batch={batch}, status={status}, operations={operations}\n".format(
                        batch=batch_result["batch"],
                        status=batch_result["status"],
                        operations=batch_result["operations"],
                    )
                )
            log_file.write("-" * 80 + "\n")

    def execute(self):
        batch_files = self._get_batch_files()
        total_batches = len(batch_files)
        print("Total prepared batch files found: " + str(total_batches))

        start_batch = self._prompt_batch_start(total_batches)
        max_batches = total_batches - start_batch + 1
        batch_count = self._prompt_batch_count(max_batches)

        selected_files = batch_files[start_batch - 1:start_batch - 1 + batch_count]
        selected_operations = 0
        payloads = []

        for batch_file in selected_files:
            payload = json.loads(batch_file.read_text(encoding="utf-8"))
            payloads.append((batch_file, payload))
            selected_operations += len(payload.get("Operations", []))

        if not self._prompt_confirmation(selected_files, selected_operations):
            print("Execution cancelled. No delete requests were submitted.")
            return

        batch_results = []
        for batch_file, payload in payloads:
            response = self.session.post(
                self.idcs_url + "/admin/v1/Bulk",
                headers=self.bulk_headers,
                params={"forceDelete": True},
                data=json.dumps(payload),
                timeout=self.REQUEST_TIMEOUT,
            )
            batch_results.append(
                {
                    "batch": batch_file.name,
                    "status": response.status_code,
                    "operations": len(payload.get("Operations", [])),
                }
            )
            print(
                "Executed {batch} with status {status}".format(
                    batch=batch_file.name,
                    status=response.status_code,
                )
            )

        self._write_execution_log(selected_files, batch_results)
        print("Execution complete. Details written to " + str(self.LOG_FILE))


executor = BatchDeleteExecutor()
executor.execute()
