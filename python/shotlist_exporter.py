import os


def create_shotlist(shotgun, entity_type, entity_ids, app):

    app.log_info("=== Shotlist Export Started ===")
    app.log_info(f"Entity type: {entity_type}")
    app.log_info(f"Entity IDs: {entity_ids}")

    shots = shotgun.find(
        "Shot",
        [["id", "in", entity_ids]],
        ["code"]
    )

    for shot in shots:
        app.log_info(f"Found shot: {shot['code']}")

    return None