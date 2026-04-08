#!/usr/local/autopkg/python

from __future__ import absolute_import

import json
import subprocess

from autopkglib import Processor, ProcessorError

__all__ = ["TeamsAdaptiveCardNotifier"]


class TeamsAdaptiveCardNotifier(Processor):
    description = "Posts an Adaptive Card to a Teams workflow webhook."
    input_variables = {
        "teams_workflow_webhook_url": {
            "required": True,
            "description": "Teams workflow webhook URL.",
        },
        "munki_repo_changed": {
            "required": False,
            "description": "Whether the Munki repo changed.",
        },
        "munki_importer_summary_result": {
            "required": False,
            "description": "Summary result from MunkiImporter.",
        },
        "NAME": {
            "required": False,
            "description": "Package name.",
        },
        "FAIL_ON_NOTIFICATION_ERROR": {
            "required": False,
            "description": "Fail recipe if notification fails. Defaults to False.",
        },
    }
    output_variables = {}

    def _bool_value(self, value):
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().lower() in ["true", "yes", "1"]
        return bool(value)

    def _build_card(self):
        summary = self.env.get("munki_importer_summary_result", {}) or {}
        data = summary.get("data", {}) or {}

        name = data.get("name") or self.env.get("NAME", "Unknown package")
        version = data.get("version", "Unknown")
        catalogs = data.get("catalogs", "")
        pkginfo_path = data.get("pkginfo_path", "")
        pkg_repo_path = data.get("pkg_repo_path", "")
        repo_changed = str(
            self._bool_value(self.env.get("munki_repo_changed", False))
        ).lower()

        return {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "AutoPkg update published",
                    "weight": "Bolder",
                    "size": "Large",
                    "wrap": True,
                },
                {
                    "type": "TextBlock",
                    "text": name,
                    "weight": "Bolder",
                    "size": "Medium",
                    "wrap": True,
                    "spacing": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": "Version %s" % version,
                    "isSubtle": True,
                    "wrap": True,
                },
                {
                    "type": "FactSet",
                    "facts": [
                        {"title": "Status", "value": "imported"},
                        {"title": "Catalogues", "value": catalogs},
                        {"title": "Repo changed", "value": repo_changed},
                    ],
                },
                {
                    "type": "TextBlock",
                    "text": "**Pkginfo**  \n%s" % pkginfo_path,
                    "wrap": True,
                    "spacing": "Medium",
                },
                {
                    "type": "TextBlock",
                    "text": "**Package**  \n%s" % pkg_repo_path,
                    "wrap": True,
                },
            ],
        }

    def _post_json(self, webhook_url, payload):
        payload_json = json.dumps(payload)

        cmd = [
            "/usr/bin/curl",
            "--silent",
            "--show-error",
            "--fail",
            "--location",
            "-X", "POST",
            "-H", "Content-Type: application/json",
            "-d", payload_json,
            webhook_url,
        ]

        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        stdout, stderr = proc.communicate()

        stdout = stdout.decode("utf-8", errors="replace").strip()
        stderr = stderr.decode("utf-8", errors="replace").strip()

        if proc.returncode != 0:
            raise ProcessorError(
                "Failed posting Teams Adaptive Card. "
                "returncode: %s stdout: %s stderr: %s"
                % (proc.returncode, stdout, stderr)
            )

        self.output("Teams Adaptive Card sent successfully.")
        if stdout:
            self.output("Response: %s" % stdout)

    def main(self):
        webhook_url = self.env.get("teams_workflow_webhook_url")
        if not webhook_url:
            raise ProcessorError("No teams_workflow_webhook_url supplied.")

        if not self._bool_value(self.env.get("munki_repo_changed", False)):
            self.output("munki_repo_changed is False. Skipping notification.")
            return

        payload = self._build_card()
        self.output("Payload: %s" % json.dumps(payload, indent=2, sort_keys=True))

        fail_on_error = self._bool_value(
            self.env.get("FAIL_ON_NOTIFICATION_ERROR", False)
        )

        try:
            self._post_json(webhook_url, payload)
        except ProcessorError as err:
            if fail_on_error:
                raise
            self.output("Notification failed but recipe will continue: %s" % err)


if __name__ == "__main__":
    PROCESSOR = TeamsAdaptiveCardNotifier()
    PROCESSOR.execute_shell()
