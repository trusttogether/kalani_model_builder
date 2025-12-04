import os
import sys
import subprocess
from pathlib import Path

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_from_directory,
    url_for,
)
import yaml

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "kalani_config.yaml"
SCRIPT_PATH = BASE_DIR / "rebuild_kalani.py"
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = "kalani-automation"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("r", encoding="utf-8") as fh:
            return yaml.safe_load(fh) or {}
    return {}


def save_config(config: dict) -> None:
    with CONFIG_PATH.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(config, fh, sort_keys=False)


def ensure_section(cfg: dict, *keys):
    cursor = cfg
    for key in keys:
        if key not in cursor or cursor[key] is None:
            cursor[key] = {}
        cursor = cursor[key]
    return cursor


def parse_value(raw: str | None):
    if raw is None:
        return ""
    text = raw.strip()
    if text == "":
        return ""
    try:
        if "." in text:
            return float(text)
        return int(text)
    except ValueError:
        return text


@app.route("/", methods=["GET", "POST"])
def index():
    cfg = load_config()
    project = ensure_section(cfg, "project")
    summary_inputs = ensure_section(cfg, "summary", "inputs")
    budget_cfg = cfg.setdefault("budget", {})
    cell_inputs = cfg.setdefault("cell_inputs", {})
    row_ranges = cfg.setdefault("row_ranges", [])

    if request.method == "POST":
        uploaded = request.files.get("template_file")
        if uploaded and uploaded.filename:
            save_path = UPLOAD_DIR / uploaded.filename
            uploaded.save(save_path)
            project["template"] = str(save_path)
        else:
            template_path = request.form.get("template_path")
            if template_path:
                project["template"] = template_path.strip()

        output_path = request.form.get("output_path")
        if output_path:
            project["output"] = output_path.strip()

        summary_fields = [
            "property_name",
            "address",
            "city_state_zip",
            "room_count",
            "gross_square_feet",
            "analysis_start_month",
            "analysis_start_year",
            "operations_start_month",
            "operations_start_year",
            "hold_period_years",
            "exit_cap_rate",
            "sale_cost_rate",
            "mezz_ltc",
        ]
        for field in summary_fields:
            summary_inputs[field] = parse_value(request.form.get(field))

        for key, value in request.form.items():
            if key.startswith("budget::"):
                _, group, label = key.split("::", 2)
                entry = budget_cfg.setdefault(group, {})
                entry[label] = parse_value(value)
            elif key.startswith("cell::"):
                _, sheet, cell = key.split("::", 2)
                sheet_map = cell_inputs.setdefault(sheet, {})
                sheet_map[cell] = parse_value(value)

        for idx, rr in enumerate(row_ranges):
            sheet = request.form.get(f"row::{idx}::sheet")
            if sheet:
                rr["sheet"] = sheet.strip()
            row_val = request.form.get(f"row::{idx}::row")
            if row_val:
                try:
                    rr["row"] = int(row_val)
                except ValueError:
                    rr["row"] = row_val
            start_col = request.form.get(f"row::{idx}::start_col")
            end_col = request.form.get(f"row::{idx}::end_col")
            if start_col:
                rr["start_col"] = start_col.strip()
            if end_col:
                rr["end_col"] = end_col.strip()
            val = request.form.get(f"row::{idx}::value")
            if val is not None:
                rr["value"] = parse_value(val)

        save_config(cfg)

        try:
            subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT_PATH),
                    "--template",
                    project.get("template", ""),
                    "--output",
                    project.get("output", "updated_kalani_model_v2.xlsx"),
                ],
                check=True,
            )
            flash("Workbook regenerated successfully!", "success")
        except subprocess.CalledProcessError as exc:
            flash(f"Error running script: {exc}", "danger")

        return redirect(url_for("index"))

    return render_template(
        "index.html",
        project=project,
        summary=summary_inputs,
        budget=budget_cfg,
        cell_inputs=cell_inputs,
        row_ranges=row_ranges,
        has_output=Path(project.get("output", "updated_kalani_model_v2.xlsx")).exists(),
    )


@app.route("/download")
def download_output():
    cfg = load_config()
    project = cfg.get("project", {})
    output_path = Path(project.get("output", "updated_kalani_model_v2.xlsx"))
    if output_path.exists():
        return send_from_directory(output_path.parent, output_path.name, as_attachment=True)
    flash("Output file not found. Generate the workbook first.", "warning")
    return redirect(url_for("index"))


@app.route("/download-config")
def download_config():
    if CONFIG_PATH.exists():
        return send_from_directory(CONFIG_PATH.parent, CONFIG_PATH.name, as_attachment=True)
    flash("Config file not found.", "warning")
    return redirect(url_for("index"))


if __name__ == "__main__":
    ssl_context = None
    if os.environ.get("USE_HTTPS", "").lower() in {"1", "true", "yes"}:
        ssl_context = "adhoc"
    app.run(host="0.0.0.0", port=5000, debug=True, ssl_context=ssl_context)
