"""Build a standard-connector Power Automate package from an exported connector seed.

The seed is read as JSON data, never executed. Personal configuration stays in the
locally generated package; --public produces an intentionally unconfigured sample.
"""
from __future__ import annotations

import argparse
import copy
import json
import re
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

CONNECTORS = {"onedriveforbusiness", "excelonlinebusiness", "office365"}


def after(name: str, statuses: list[str] | None = None) -> dict:
    return {name: statuses or ["Succeeded"]} if name else {}


def condition(expression: str, yes: dict, no: dict | None = None, dependency: str = "") -> dict:
    return {"type": "If", "expression": {"and": [{"equals": [expression, True]}]},
            "actions": yes, "else": {"actions": no or {}}, "runAfter": after(dependency)}


def flow_definition(config: dict, authenticators: dict) -> dict:
    source = config.get("SourceId", "portsmouth-2026-2027")
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,100}", source):
        raise ValueError("SourceId must contain only letters, digits, underscores and hyphens.")
    config = {**config, "SourceId": source, "ApplyChanges": False}
    config.setdefault("Rules", {})
    feed_path = "@concat('/timetable-sync-feed-', variables('Run_id'), '.ics')"
    state_path = "@concat('/timetable-sync-state-', variables('Config')['SourceId'], '.json')"
    local_start = "concat(formatDateTime(convertTimeZone(variables('Plan_time'),'UTC','GMT Standard Time'),'yyyy-MM-dd'),'T00:00:00')"

    def api(connection: str, operation: str, params: dict, dependency: str = "", no_retry: bool = False) -> dict:
        action = {"type": "OpenApiConnection", "inputs": {"parameters": params, "host": {
            "apiId": "/providers/Microsoft.PowerApps/apis/shared_" + connection,
            "connectionName": "shared_" + connection, "operationId": operation},
            "authentication": copy.deepcopy(authenticators[connection])}, "runAfter": after(dependency),
            "runtimeConfiguration": {"secureData": {"properties": ["inputs", "outputs"]}}}
        if no_retry:
            action["inputs"]["retryPolicy"] = {"type": "none"}
        return action

    def set_var(name: str, value: object, dependency: str = "") -> dict:
        return {"type": "SetVariable", "inputs": {"name": name, "value": value}, "runAfter": after(dependency)}

    def script(mode: str, extra: dict | None = None, dependency: str = "") -> dict:
        parameters = {
            "source": "@variables('Config')['WorkbookSource']", "drive": "@variables('Config')['WorkbookDrive']",
            "file": "@variables('Config')['WorkbookFile']", "scriptId": "@variables('Config')['ScriptId']",
            "ScriptParameters/mode": mode, "ScriptParameters/icsContent": "unused",
            "ScriptParameters/existingEventsJson": "[]", "ScriptParameters/stateJson": "@variables('State_text')",
            "ScriptParameters/sourceId": "@variables('Config')['SourceId']", "ScriptParameters/feedRunId": "@variables('Run_id')",
            "ScriptParameters/fetchedAtUtc": "@variables('Plan_time')", "ScriptParameters/scanComplete": False,
            "ScriptParameters/planJson": "{}", "ScriptParameters/acknowledgementsJson": "[]",
            "ScriptParameters/rulesJson": "@string(variables('Config')['Rules'])", "ScriptParameters/nowUtc": "@utcNow()"}
        parameters.update(extra or {})
        # Dynamic script IDs cannot expose flattened designer fields at import.
        # Supply the required parameter object accepted by RunScriptProd.
        parameters["ScriptParameters"] = {
            key.split("/", 1)[1]: parameters.pop(key)
            for key in list(parameters) if key.startswith("ScriptParameters/")
        }
        return api("excelonlinebusiness", "RunScriptProd", parameters, dependency)

    def stop(code: str) -> dict:
        return {"type": "Terminate", "inputs": {"runStatus": "Failed", "runError": {
            "code": code, "message": "Timetable sync stopped. Review the Preview worksheet and flow history. No additional operations will be attempted."}}}

    def ack(kind: str, loop: str, identifier: str, dependency: str) -> dict:
        return {"type": "AppendToArrayVariable", "inputs": {"name": "Acknowledgements", "value": {
            "kind": kind, "key": "@items('" + loop + "')['key']", "id": identifier}}, "runAfter": after(dependency)}

    def safe_latest(read_action: str, loop: str) -> str:
        prefix = "concat('Timetable Sync [v1:',uriComponent(variables('Config')['SourceId']),':',items('" + loop + "')['key'],']')"
        return "@and(contains(coalesce(body('" + read_action + "')?['body'],'')," + prefix + "),equals(body('" + read_action + "')?['isAllDay'],false),empty(body('" + read_action + "')?['seriesMasterId']),or(empty(body('" + read_action + "')?['recurrence']),equals(toLower(string(body('" + read_action + "')?['recurrence'])),'none')),empty(concat(coalesce(body('" + read_action + "')?['requiredAttendees'],''),coalesce(body('" + read_action + "')?['optionalAttendees'],''),coalesce(body('" + read_action + "')?['resourceAttendees'],''))))"

    variables = [
        {"name": "Config", "type": "object", "value": config}, {"name": "Run_id", "type": "string", "value": "@guid()"},
        {"name": "Run_started", "type": "string", "value": "@utcNow()"}, {"name": "Plan_time", "type": "string", "value": "@utcNow()"},
        {"name": "State_text", "type": "string", "value": "{}"}, {"name": "Outlook_events", "type": "array", "value": []},
        {"name": "Page_count", "type": "integer", "value": 256}, {"name": "Skip_count", "type": "integer", "value": 0},
        {"name": "Acknowledgements", "type": "array", "value": []}, {"name": "Failed", "type": "boolean", "value": False}]
    actions = {}
    previous_initializer = ""
    for index, variable in enumerate(variables):
        name = "Initialize" if index == 0 else "Initialize_" + variable["name"]
        actions[name] = {"type": "InitializeVariable", "inputs": {"variables": [variable]}, "runAfter": after(previous_initializer)}
        if index == 0:
            actions[name]["runtimeConfiguration"] = {"secureData": {"properties": ["inputs", "outputs"]}}
        previous_initializer = name

    read_page = api("office365", "GetEventsCalendarViewV3", {
        "calendarId": "@variables('Config')['CalendarId']", "startDateTimeUtc": "@convertToUtc(" + local_start + ",'GMT Standard Time')",
        "endDateTimeUtc": "@convertToUtc(addDays(" + local_start + ",365),'GMT Standard Time')", "$top": 256, "$skip": "@variables('Skip_count')"})
    page_actions = {"Read_page": read_page,
                    "Merge_page": {"type": "Compose", "inputs": "@union(variables('Outlook_events'),body('Read_page')['value'])", "runAfter": after("Read_page")},
                    "Save_page": set_var("Outlook_events", "@outputs('Merge_page')", "Merge_page"),
                    "Count_page": set_var("Page_count", "@length(body('Read_page')['value'])", "Save_page"),
                    "Advance_page": {"type": "IncrementVariable", "inputs": {"name": "Skip_count", "value": 256}, "runAfter": after("Count_page")}}
    normalize = {"id": "@item()['id']", "subject": "@coalesce(item()?['subject'],'')", "start": "@coalesce(item()?['startWithTimeZone'],item()?['start'])",
                 "end": "@coalesce(item()?['endWithTimeZone'],item()?['end'])", "timeZone": "@coalesce(item()?['timeZone'],'')",
                 "body": "@coalesce(item()?['body'],'')", "location": "@coalesce(item()?['location'],'')", "showAs": "@coalesce(item()?['showAs'],'unknown')",
                 "hasAttendees": "@not(empty(concat(coalesce(item()?['requiredAttendees'],''),coalesce(item()?['optionalAttendees'],''),coalesce(item()?['resourceAttendees'],''))))",
                 "isAllDay": "@coalesce(item()?['isAllDay'],false)", "isRecurring": "@or(not(empty(item()?['seriesMasterId'])),and(not(empty(item()?['recurrence'])),not(equals(toLower(string(item()?['recurrence'])),'none'))))",
                 "categories": "@coalesce(item()?['categories'],json('[]'))", "importance": "@coalesce(item()?['importance'],'normal')",
                 "isReminderOn": "@coalesce(item()?['isReminderOn'],true)", "reminderMinutesBeforeStart": "@coalesce(item()?['reminderMinutesBeforeStart'],15)"}
    prepare = {
        "Download_feed": api("onedriveforbusiness", "CopyFile", {"source": "@variables('Config')['FeedUrl']", "destination": feed_path, "overwrite": False}),
        "Wait_for_download": {"type": "Wait", "inputs": {"interval": {"count": 5, "unit": "Second"}}, "runAfter": after("Download_feed")},
        "Feed_metadata": api("onedriveforbusiness", "GetFileMetadataByPath", {"path": feed_path}, "Wait_for_download"),
        "Verify_new_file": condition("@greaterOrEquals(ticks(body('Feed_metadata')['LastModified']),sub(ticks(variables('Run_started')),300000000))", {}, {"Stop_stale_file": stop("StaleFeed")}, "Feed_metadata"),
        "Feed_content": api("onedriveforbusiness", "GetFileContent", {"id": "@body('Feed_metadata')['Id']", "inferContentType": False}, "Verify_new_file"),
        "Try_state": {"type": "Scope", "actions": {
            "State_metadata": api("onedriveforbusiness", "GetFileMetadataByPath", {"path": state_path}),
            "State_content": api("onedriveforbusiness", "GetFileContent", {"id": "@body('State_metadata')['Id']", "inferContentType": False}, "State_metadata")}, "runAfter": after("Feed_content")},
        "Handle_state": condition("@equals(actions('Try_state')['status'],'Succeeded')", {
            "Read_state_text": set_var("State_text", "@base64ToString(body('State_content')['$content'])")}, {
            "Check_missing_state": condition("@or(equals(outputs('State_metadata')?['statusCode'],404),equals(outputs('State_content')?['statusCode'],404))", {}, {"Stop_unreadable_state": stop("UnreadableState")})}),
        "Set_plan_time": set_var("Plan_time", "@utcNow()", "Handle_state"),
        "Scan_calendar": {"type": "Until", "expression": "@less(variables('Page_count'),256)", "limit": {"count": 20, "timeout": "PT5M"}, "actions": page_actions, "runAfter": after("Set_plan_time")},
        "Normalize_calendar": {"type": "Select", "inputs": {"from": "@variables('Outlook_events')", "select": normalize}, "runAfter": after("Scan_calendar")},
        "Run_plan": script("plan", {"ScriptParameters/icsContent": "@base64ToString(body('Feed_content')['$content'])",
            "ScriptParameters/existingEventsJson": "@string(body('Normalize_calendar'))", "ScriptParameters/fetchedAtUtc": "@body('Feed_metadata')['LastModified']",
            "ScriptParameters/scanComplete": "@and(less(variables('Page_count'),256),lessOrEquals(length(variables('Outlook_events')),5000))",
            "ScriptParameters/nowUtc": "@variables('Plan_time')"}, "Normalize_calendar")}
    prepare["Handle_state"]["runAfter"] = after("Try_state", ["Succeeded", "Failed", "TimedOut"])
    actions["Prepare_and_plan"] = {"type": "Scope", "actions": prepare, "runAfter": after(previous_initializer)}

    field_names = ["subject", "start", "end", "timeZone", "body", "location", "showAs", "categories", "importance", "isReminderOn", "reminderMinutesBeforeStart", "isAllDay"]
    create_params = {"table": "@variables('Config')['CalendarId']"}
    create_params.update({"item/" + field: "@items('Create_classes')['fields']['" + field + "']" for field in field_names})
    create_params.update({"item/recurrence": "none", "item/responseRequested": False})
    create_loop = {"type": "Foreach", "foreach": "@body('Run_plan')['result']['creates']", "actions": {
        "Create_class": api("office365", "V4CalendarPostItem", create_params, no_retry=True),
        "Ack_create": ack("create", "Create_classes", "@body('Create_class')['id']", "Create_class")},
        "runtimeConfiguration": {"concurrency": {"repetitions": 1}}, "runAfter": {}}
    create_loop["actions"]["Flag_create_failure"] = set_var("Failed", True)
    create_loop["actions"]["Flag_create_failure"]["runAfter"] = after("Create_class", ["Failed", "TimedOut"])
    create_loop["actions"] = {"Create_if_healthy": condition("@equals(variables('Failed'),false)", create_loop["actions"])}
    update_params = {"table": "@variables('Config')['CalendarId']", "id": "@items('Update_classes')['id']"}
    update_params.update({"item/" + field: "@items('Update_classes')['fields']['" + field + "']" for field in field_names})
    update_params.update({"item/recurrence": "none", "item/responseRequested": False,
        "item/showAs": "@if(equals(body('Read_update_target')['showAs'],items('Update_classes')['observedShowAs']),items('Update_classes')['fields']['showAs'],body('Read_update_target')['showAs'])",
        "item/categories": "@if(empty(trim(coalesce(variables('Config')?['Rules']?['calendarCategory'],''))),coalesce(body('Read_update_target')?['categories'],json('[]')),union(coalesce(body('Read_update_target')?['categories'],json('[]')),createArray(trim(variables('Config')['Rules']['calendarCategory']))))",
        "item/importance": "@coalesce(body('Read_update_target')?['importance'],'normal')",
        "item/isReminderOn": "@coalesce(body('Read_update_target')?['isReminderOn'],true)",
        "item/reminderMinutesBeforeStart": "@coalesce(body('Read_update_target')?['reminderMinutesBeforeStart'],15)"})
    update_loop = {"type": "Foreach", "foreach": "@body('Run_plan')['result']['updates']", "actions": {
        "Read_update_target": api("office365", "V3CalendarGetItem", {"table": "@variables('Config')['CalendarId']", "id": "@items('Update_classes')['id']"}),
        "Check_update_ownership": condition(safe_latest("Read_update_target", "Update_classes"), {
            "Update_class": api("office365", "V4CalendarPatchItem", update_params, no_retry=True),
            "Ack_update": ack("update", "Update_classes", "@items('Update_classes')['id']", "Update_class")}, {"Flag_changed_update_target": set_var("Failed", True)}, "Read_update_target")},
        "runtimeConfiguration": {"concurrency": {"repetitions": 1}}, "runAfter": after("Create_classes")}
    update_loop["actions"]["Flag_update_read_failure"] = set_var("Failed", True)
    update_loop["actions"]["Flag_update_read_failure"]["runAfter"] = after("Read_update_target", ["Failed", "TimedOut"])
    update_write = update_loop["actions"]["Check_update_ownership"]["actions"]
    update_write["Flag_update_failure"] = set_var("Failed", True)
    update_write["Flag_update_failure"]["runAfter"] = after("Update_class", ["Failed", "TimedOut"])
    update_loop["actions"] = {"Update_if_healthy": condition("@equals(variables('Failed'),false)", update_loop["actions"])}
    delete_loop = {"type": "Foreach", "foreach": "@body('Run_plan')['result']['deletes']", "actions": {
        "Read_delete_target": api("office365", "V3CalendarGetItem", {"table": "@variables('Config')['CalendarId']", "id": "@items('Delete_classes')['id']"}),
        "Check_delete_ownership": condition(safe_latest("Read_delete_target", "Delete_classes"), {
            "Delete_class": api("office365", "CalendarDeleteItem_V2", {"calendar": "@variables('Config')['CalendarId']", "event": "@items('Delete_classes')['id']"}, no_retry=True),
            "Ack_delete": ack("delete", "Delete_classes", "@items('Delete_classes')['id']", "Delete_class")}, {"Flag_changed_delete_target": set_var("Failed", True)}, "Read_delete_target")},
        "runtimeConfiguration": {"concurrency": {"repetitions": 1}}, "runAfter": after("Update_classes")}
    delete_loop["actions"]["Flag_delete_read_failure"] = set_var("Failed", True)
    delete_loop["actions"]["Flag_delete_read_failure"]["runAfter"] = after("Read_delete_target", ["Failed", "TimedOut"])
    delete_write = delete_loop["actions"]["Check_delete_ownership"]["actions"]
    delete_write["Flag_delete_failure"] = set_var("Failed", True)
    delete_write["Flag_delete_failure"]["runAfter"] = after("Delete_class", ["Failed", "TimedOut"])
    delete_loop["actions"] = {"Delete_if_healthy": condition("@equals(variables('Failed'),false)", delete_loop["actions"])}
    apply_actions = {"Create_classes": create_loop, "Update_classes": update_loop, "Delete_classes": delete_loop,
        "Commit_state": script("commit", {"ScriptParameters/planJson": "@body('Run_plan')['result']['planJson']", "ScriptParameters/acknowledgementsJson": "@string(variables('Acknowledgements'))"}, "Delete_classes"),
        "Persist_state": condition("@not(empty(body('State_metadata')?['Id']))", {
            "Update_state_file": api("onedriveforbusiness", "UpdateFile", {"id": "@body('State_metadata')['Id']", "body": "@body('Commit_state')['result']['stateJson']"})}, {
            "Create_state_file": api("onedriveforbusiness", "CreateFile", {"folderPath": "/", "name": "@concat('timetable-sync-state-',variables('Config')['SourceId'],'.json')", "body": "@body('Commit_state')['result']['stateJson']"})}, "Commit_state")}
    checkpoint = {"Commit_state": apply_actions.pop("Commit_state"), "Persist_state": apply_actions.pop("Persist_state")}
    checkpoint["Commit_state"]["runAfter"] = {}
    apply_actions["Commit_if_healthy"] = condition("@equals(variables('Failed'),false)", checkpoint, dependency="Delete_classes")
    actions["Apply_or_preview"] = condition("@variables('Config')['ApplyChanges']", apply_actions, {
        "Preview_only": {"type": "Compose", "inputs": "Preview generated in the workbook. No calendar changes applied."}}, "Prepare_and_plan")
    for name, dependency, reason in [("Source_failure", "Prepare_and_plan", "source"), ("Apply_failure", "Apply_or_preview", "apply")]:
        actions[name] = {"type": "Scope", "actions": {"Flag_" + name: set_var("Failed", True),
            "Report_" + name: script("failure", {"ScriptParameters/planJson": reason}, "Flag_" + name)},
            "runAfter": after(dependency, ["Failed", "TimedOut"])}
    apply_failure = actions["Apply_failure"]
    apply_failure["actions"] = {"Report_if_apply_failed": condition("@or(equals(variables('Failed'),true),not(equals(actions('Apply_or_preview')['status'],'Succeeded')))", apply_failure["actions"])}
    apply_failure["runAfter"] = after("Apply_or_preview", ["Succeeded", "Failed", "TimedOut"])
    actions["Cleanup_feed"] = condition("@not(empty(body('Feed_metadata')?['Id']))", {
        "Remove_temporary_feed": api("onedriveforbusiness", "DeleteFile", {"id": "@body('Feed_metadata')['Id']"})})
    actions["Cleanup_feed"]["runAfter"] = {name: ["Succeeded", "Failed", "Skipped", "TimedOut"] for name in ["Source_failure", "Apply_failure", "Apply_or_preview"]}
    actions["Finish"] = condition("@and(equals(variables('Failed'),false),equals(actions('Cleanup_feed')['status'],'Succeeded'))", {
        "Summary": {"type": "Compose", "inputs": {"applied": "@variables('Config')['ApplyChanges']", "sessions": "@body('Run_plan')['result']['sessionCount']",
            "creates": "@length(body('Run_plan')['result']['creates'])", "updates": "@length(body('Run_plan')['result']['updates'])",
            "deletes": "@length(body('Run_plan')['result']['deletes'])", "held": "@length(body('Run_plan')['result']['held'])"}}}, {"Stop_failed_run": stop("SyncIncomplete")})
    actions["Finish"]["runAfter"] = after("Cleanup_feed", ["Succeeded", "Failed", "TimedOut"])
    return {"$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
            "contentVersion": "1.0.0.0", "parameters": {"$authentication": {"defaultValue": {}, "type": "SecureObject"}, "$connections": {"defaultValue": {}, "type": "Object"}},
            "triggers": {"manual": {"type": "Request", "kind": "Button", "inputs": {"schema": {"type": "object", "properties": {}, "required": []}},
                                    "runtimeConfiguration": {"concurrency": {"runs": 1}}}}, "actions": actions, "outputs": {}}


def load_seed(filename: Path) -> tuple[dict, dict, dict, dict]:
    with zipfile.ZipFile(filename) as archive:
        definition_path = next(name for name in archive.namelist() if name.endswith("/definition.json"))
        document = json.loads(archive.read(definition_path))
        root = definition_path.rsplit("/", 1)[0]
        return document, json.loads(archive.read("manifest.json")), json.loads(archive.read(root + "/apisMap.json")), json.loads(archive.read(root + "/connectionsMap.json"))


def extract_config(document: dict) -> dict:
    actions = document["properties"]["definition"]["actions"]
    by_operation = {action["inputs"]["host"]["operationId"]: action for action in actions.values() if action.get("type") == "OpenApiConnection"}
    script = by_operation["RunScriptProd"]["inputs"]["parameters"]
    reader = by_operation.get("GetEventsCalendarViewV3", {}).get("inputs", {}).get("parameters", {})
    return {"FeedUrl": by_operation["CopyFile"]["inputs"]["parameters"]["source"], "SourceId": "portsmouth-2026-2027",
            "WorkbookSource": script["source"], "WorkbookDrive": script["drive"], "WorkbookFile": script["file"], "ScriptId": script["scriptId"],
            "CalendarId": reader.get("calendarId", "CONFIGURE_MAIN_CALENDAR_ID"), "Rules": {}}


def build_package(seed: Path, output: Path, config: dict, public: bool = False) -> dict:
    document, manifest, api_map, connection_map = load_seed(seed)
    authenticators: dict = {}
    for action in document["properties"]["definition"]["actions"].values():
        if action.get("type") == "OpenApiConnection":
            name = action["inputs"]["host"]["apiId"].rsplit("shared_", 1)[1]
            authenticators[name] = copy.deepcopy(action["inputs"].get("authentication", "@parameters('$authentication')"))
    if set(authenticators) != CONNECTORS:
        raise ValueError("The exported seed must contain connected OneDrive, Excel Run script, and Outlook calendar-view actions.")
    if public:
        config = {"FeedUrl": "https://example.invalid/calendar.ics", "SourceId": "portsmouth-2026-2027", "WorkbookSource": "me",
                  "WorkbookDrive": "CONFIGURE_WORKBOOK_DRIVE", "WorkbookFile": "CONFIGURE_WORKBOOK_FILE", "ScriptId": "CONFIGURE_SCRIPT_ID",
                  "CalendarId": "CONFIGURE_MAIN_CALENDAR_ID", "Rules": {}}
    flow = flow_definition(config, authenticators)
    flow_id = str(uuid.uuid4())
    old_root = next(key for key, resource in manifest["resources"].items() if resource["type"] == "Microsoft.Flow/flows")
    manifest["resources"][flow_id] = manifest["resources"].pop(old_root)
    manifest["resources"][flow_id]["suggestedCreationType"] = "New"
    manifest["resources"][flow_id]["details"]["displayName"] = "Timetable Sync - Standard Preview"
    manifest["details"].update(displayName="Timetable Sync - Standard Preview", description="Standard connectors only. Manual preview; ApplyChanges is false. A fresh source download is required before any calendar operation.", creator="N/A", sourceEnvironment="", createdTime=datetime.now(timezone.utc).isoformat(), packageTelemetryId=str(uuid.uuid4()))
    for resource in manifest["resources"].values():
        if resource["type"] == "Microsoft.PowerApps/apis/connections":
            resource["details"]["displayName"] = "Select your school connection"
    properties = document["properties"]
    properties["definition"] = flow
    properties["displayName"] = "Timetable Sync - Standard Preview"
    properties["flowFailureAlertSubscribed"] = False
    if public:
        for reference in properties["connectionReferences"].values():
            reference["connectionName"] = ""
    document.update(name=flow_id, id="/providers/Microsoft.Flow/flows/" + flow_id)
    folder = "Microsoft.Flow/flows/" + flow_id
    payloads = {"manifest.json": manifest, "Microsoft.Flow/flows/manifest.json": {"packageSchemaVersion": "1.0", "flowAssets": {"assetPaths": [flow_id]}},
                folder + "/definition.json": document, folder + "/apisMap.json": api_map, folder + "/connectionsMap.json": connection_map}
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for filename, value in payloads.items():
            archive.writestr(filename, json.dumps(value, indent=2, ensure_ascii=True))
    return flow


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("seed", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--config", type=Path)
    parser.add_argument("--public", action="store_true")
    parser.add_argument("--extract-config", action="store_true")
    args = parser.parse_args()
    document = load_seed(args.seed)[0]
    config = json.loads(args.config.read_text(encoding="utf-8")) if args.config else extract_config(document)
    if args.extract_config:
        args.output.write_text(json.dumps(config, indent=2), encoding="utf-8")
        print("Private configuration extracted. Keep this file out of Git.")
        return
    flow = build_package(args.seed, args.output, config, args.public)
    print(f"Built {args.output.name}: {len(flow['actions'])} top-level steps, three standard connectors, manual preview only.")


if __name__ == "__main__":
    main()
