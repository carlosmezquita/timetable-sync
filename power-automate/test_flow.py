"""Safety checks for the generated cloud workflow, using synthetic export metadata."""
import importlib.util
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

spec = importlib.util.spec_from_file_location("flow_builder", Path(__file__).with_name("build-flow.py"))
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)

CONFIG = {"FeedUrl": "https://example.invalid/synthetic-feed.ics", "SourceId": "synthetic-school", "WorkbookSource": "me",
          "WorkbookDrive": "synthetic-drive", "WorkbookFile": "synthetic-file", "ScriptId": "synthetic-script", "CalendarId": "synthetic-main-calendar", "Rules": {}, "ApplyChanges": True}
AUTH = {name: {"type": "Raw", "value": "synthetic-auth-expression"} for name in builder.CONNECTORS}


def walk(actions, ancestors=()):
    for name, action in actions.items():
        yield ancestors + (name,), action
        yield from walk(action.get("actions", {}), ancestors + (name, "actions"))
        yield from walk(action.get("else", {}).get("actions", {}), ancestors + (name, "else"))


def seed_archive(path):
    flow_id = "synthetic-export-flow"
    resources = {flow_id: {"type": "Microsoft.Flow/flows", "suggestedCreationType": "Update", "details": {"displayName": "Seed"}, "dependsOn": []}}
    api_map = {}
    conn_map = {}
    references = {}
    actions = {}
    operations = {"onedriveforbusiness": "CopyFile", "excelonlinebusiness": "RunScriptProd", "office365": "GetEventsCalendarViewV3"}
    for connector, operation in operations.items():
        api_id = connector + "-api"
        conn_id = connector + "-connection"
        resources[api_id] = {"type": "Microsoft.PowerApps/apis", "name": "shared_" + connector}
        resources[conn_id] = {"type": "Microsoft.PowerApps/apis/connections", "details": {"displayName": "PRIVATE_SYNTHETIC_ACCOUNT"}, "dependsOn": [api_id]}
        resources[flow_id]["dependsOn"] += [api_id, conn_id]
        api_map["shared_" + connector] = api_id
        conn_map["shared_" + connector] = conn_id
        references["shared_" + connector] = {"connectionName": "PRIVATE_SYNTHETIC_CONNECTION", "source": "Invoker", "id": "/providers/Microsoft.PowerApps/apis/shared_" + connector}
        actions[connector] = {"type": "OpenApiConnection", "inputs": {"host": {"apiId": "/providers/Microsoft.PowerApps/apis/shared_" + connector, "operationId": operation}, "authentication": AUTH[connector], "parameters": {}}}
    document = {"name": flow_id, "id": "synthetic", "type": "Microsoft.Flow/flows", "properties": {
        "displayName": "Seed", "definition": {"actions": actions}, "connectionReferences": references, "apiId": "synthetic-api"}}
    manifest = {"schema": "1.0", "details": {}, "resources": resources}
    with zipfile.ZipFile(path, "w") as archive:
        for filename, data in {"manifest.json": manifest, f"Microsoft.Flow/flows/{flow_id}/definition.json": document,
                               f"Microsoft.Flow/flows/{flow_id}/apisMap.json": api_map, f"Microsoft.Flow/flows/{flow_id}/connectionsMap.json": conn_map}.items():
            archive.writestr(filename, json.dumps(data))


class WorkflowSafety(unittest.TestCase):
    def setUp(self):
        self.flow = builder.flow_definition(CONFIG, AUTH)

    def test_calendar_and_checkpoint_writes_only_live_inside_apply_gate(self):
        mutators = {"V4CalendarPostItem", "V4CalendarPatchItem", "CalendarDeleteItem_V2", "UpdateFile", "CreateFile"}
        found = []
        for ancestry, action in walk(self.flow["actions"]):
            inputs = action.get("inputs", {})
            op = inputs.get("host", {}).get("operationId") if isinstance(inputs, dict) else None
            if op in mutators:
                self.assertEqual(ancestry[:2], ("Apply_or_preview", "actions"))
                found.append(op)
        self.assertEqual(set(found), mutators)
        values = self.flow["actions"]["Initialize"]["inputs"]["variables"]
        self.assertIs(next(v for v in values if v["name"] == "Config")["value"]["ApplyChanges"], False)

    def test_source_is_unique_verified_and_required_before_planning_or_apply(self):
        prepare = self.flow["actions"]["Prepare_and_plan"]["actions"]
        params = prepare["Download_feed"]["inputs"]["parameters"]
        self.assertIs(params["overwrite"], False)
        self.assertIn("Run_id", params["destination"])
        self.assertEqual(prepare["Feed_content"]["runAfter"], {"Verify_new_file": ["Succeeded"]})
        self.assertEqual(self.flow["actions"]["Apply_or_preview"]["runAfter"], {"Prepare_and_plan": ["Succeeded"]})
        self.assertIn("LastModified", prepare["Verify_new_file"]["expression"]["and"][0]["equals"][0])

    def test_no_premium_network_action_or_custom_connector(self):
        for _, action in walk(self.flow["actions"]):
            self.assertNotIn(action["type"], {"Http", "HttpWebhook", "HttpWithSwagger"})
            inputs = action.get("inputs", {})
            host = inputs.get("host") if isinstance(inputs, dict) else None
            if host:
                self.assertIn(host["apiId"].rsplit("shared_", 1)[1], builder.CONNECTORS)

    def test_pagination_has_an_explicit_end_and_incomplete_scan_guard(self):
        prepare = self.flow["actions"]["Prepare_and_plan"]["actions"]
        scan = prepare["Scan_calendar"]
        self.assertEqual(scan["limit"], {"count": 20, "timeout": "PT5M"})
        self.assertEqual(scan["actions"]["Read_page"]["inputs"]["parameters"]["$top"], 256)
        self.assertIn("less(variables('Page_count'),256)", prepare["Run_plan"]["inputs"]["parameters"]["ScriptParameters"]["scanComplete"])

    def test_mutations_are_serial_and_create_retries_are_disabled(self):
        apply = self.flow["actions"]["Apply_or_preview"]["actions"]
        for name in ["Create_classes", "Update_classes", "Delete_classes"]:
            self.assertEqual(apply[name]["runtimeConfiguration"]["concurrency"]["repetitions"], 1)
        self.assertEqual(apply["Create_classes"]["actions"]["Create_if_healthy"]["actions"]["Create_class"]["inputs"]["retryPolicy"], {"type": "none"})
        self.assertEqual(self.flow["triggers"]["manual"]["runtimeConfiguration"]["concurrency"]["runs"], 1)

    def test_update_and_delete_recheck_ownership_and_meeting_safety(self):
        apply = self.flow["actions"]["Apply_or_preview"]["actions"]
        for loop, check, reader in [("Update_classes", "Check_update_ownership", "Read_update_target"), ("Delete_classes", "Check_delete_ownership", "Read_delete_target")]:
            wrapper = "Update_if_healthy" if loop == "Update_classes" else "Delete_if_healthy"
            gate = apply[loop]["actions"][wrapper]["actions"][check]
            self.assertEqual(gate["runAfter"], {reader: ["Succeeded"]})
            expression = gate["expression"]["and"][0]["equals"][0]
            for token in ["Timetable Sync [v1:", "requiredAttendees", "optionalAttendees", "resourceAttendees", "isAllDay", "seriesMasterId"]:
                self.assertIn(token, expression)

    def test_last_moment_free_busy_and_reminder_edits_are_preserved(self):
        update = self.flow["actions"]["Apply_or_preview"]["actions"]["Update_classes"]["actions"]["Update_if_healthy"]["actions"]["Check_update_ownership"]["actions"]["Update_class"]["inputs"]["parameters"]
        self.assertIn("observedShowAs", update["item/showAs"])
        self.assertIn("Read_update_target", update["item/showAs"])
        self.assertIn("Read_update_target", update["item/categories"])
        self.assertIn("union(", update["item/categories"])
        self.assertIn("calendarCategory", update["item/categories"])
        self.assertNotIn("['fields']['categories']", update["item/categories"])
        self.assertIn("Read_update_target", update["item/reminderMinutesBeforeStart"])

    def test_checkpoint_is_written_only_after_complete_operation_acknowledgements(self):
        apply = self.flow["actions"]["Apply_or_preview"]["actions"]
        self.assertEqual(apply["Commit_if_healthy"]["runAfter"], {"Delete_classes": ["Succeeded"]})
        checkpoint = apply["Commit_if_healthy"]["actions"]
        self.assertIn("Acknowledgements", checkpoint["Commit_state"]["inputs"]["parameters"]["ScriptParameters"]["acknowledgementsJson"])
        self.assertEqual(checkpoint["Persist_state"]["runAfter"], {"Commit_state": ["Succeeded"]})

    def test_a_failure_stops_later_iterations_without_illegal_loop_termination(self):
        apply = self.flow["actions"]["Apply_or_preview"]["actions"]
        for loop in ["Create_classes", "Update_classes", "Delete_classes"]:
            guard = next(iter(apply[loop]["actions"].values()))
            self.assertIn("variables('Failed')", guard["expression"]["and"][0]["equals"][0])
            flags = []
            for _, child in walk(apply[loop]["actions"]):
                self.assertNotEqual(child["type"], "Terminate")
                if child["type"] == "SetVariable":
                    flags.append(child["inputs"])
            self.assertIn({"name": "Failed", "value": True}, flags)

    def test_dynamic_script_actions_supply_required_parameter_object(self):
        scripts = [action for _, action in walk(self.flow["actions"])
                   if isinstance(action.get("inputs"), dict)
                   and action["inputs"].get("host", {}).get("operationId") == "RunScriptProd"]
        self.assertEqual(len(scripts), 4)
        for action in scripts:
            parameters = action["inputs"]["parameters"]
            self.assertIsInstance(parameters["ScriptParameters"], dict)
            self.assertIn("mode", parameters["ScriptParameters"])
            self.assertIn("scanComplete", parameters["ScriptParameters"])
            self.assertFalse(any(key.startswith("ScriptParameters/") for key in parameters))

    def test_outlook_none_recurrence_is_accepted_at_scan_and_write_checks(self):
        mapping = self.flow["actions"]["Prepare_and_plan"]["actions"]["Normalize_calendar"]["inputs"]["select"]
        self.assertIn("not(equals(toLower(string(item()?['recurrence'])),'none'))", mapping["isRecurring"])
        apply = self.flow["actions"]["Apply_or_preview"]["actions"]
        for loop, wrapper, check, reader in [
            ("Update_classes", "Update_if_healthy", "Check_update_ownership", "Read_update_target"),
            ("Delete_classes", "Delete_if_healthy", "Check_delete_ownership", "Read_delete_target")
        ]:
            expression = apply[loop]["actions"][wrapper]["actions"][check]["expression"]["and"][0]["equals"][0]
            self.assertIn("or(empty(body('" + reader + "')?['recurrence']),equals(toLower(string(body('" + reader + "')?['recurrence'])),'none'))", expression)

    def test_exported_sample_contains_no_account_or_connection_names(self):
        workspace = Path.cwd().resolve()
        with tempfile.TemporaryDirectory(dir=workspace, prefix=".power-automate-test-") as folder:
            self.assertTrue(Path(folder).resolve().is_relative_to(workspace))
            seed = Path(folder) / "seed.zip"
            package = Path(folder) / "sample.zip"
            seed_archive(seed)
            builder.build_package(seed, package, CONFIG, public=True)
            with zipfile.ZipFile(package) as archive:
                content = "\n".join(archive.read(name).decode() for name in archive.namelist())
                self.assertNotIn("PRIVATE_SYNTHETIC_ACCOUNT", content)
                self.assertNotIn("PRIVATE_SYNTHETIC_CONNECTION", content)
                self.assertNotIn("synthetic-drive", content)
                self.assertNotIn("synthetic-main-calendar", content)
                manifest = json.loads(archive.read("manifest.json"))
                roots = [resource for resource in manifest["resources"].values() if resource["type"] == "Microsoft.Flow/flows"]
                self.assertEqual(roots[0]["suggestedCreationType"], "New")


if __name__ == "__main__":
    unittest.main(verbosity=2)
