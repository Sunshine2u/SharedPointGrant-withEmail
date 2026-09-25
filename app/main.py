from flask import Flask, render_template, jsonify, request
from excel_service import read_excel_data, update_excel_logs, update_permission, copy_permission, delete_permission
from sharepoint_service import process_sharepoint_permission, get_server_relative_path
from datetime import datetime
import urllib.parse
import os
import threading
import time
import uuid


app = Flask(__name__)
jobs = {}
jobs_lock = threading.Lock()
grant_delay_seconds = max(0, float(os.getenv("GRANT_DELAY_SECONDS", "1")))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['GET'])
def get_data():
    try:
        data = read_excel_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/execute', methods=['POST'])
def execute_action():
    try:
        req_data = request.json
        action = req_data.get('action')
        selected_ids = req_data.get('selected_ids', [])
        job_id = uuid.uuid4().hex
        with jobs_lock:
            jobs[job_id] = {"status": "queued", "current": 0, "total": len(selected_ids), "results": []}
        thread = threading.Thread(target=run_execute_job, args=(job_id, action, selected_ids), daemon=True)
        thread.start()
        return jsonify({"success": True, "job_id": job_id})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

def run_execute_job(job_id, action, selected_ids):
    try:
        excel_info = read_excel_data()
        main_path_url = excel_info['main_path']
        parsed_main = urllib.parse.urlparse(main_path_url)
        site_name = parsed_main.path.split('/sites/')[1].split('/')[0] if '/sites/' in parsed_main.path else ''
        site_url = f"{parsed_main.scheme}://{parsed_main.netloc}/sites/{site_name}"
        selected = {str(value) for value in selected_ids}
        items = [item for item in excel_info['permissions'] if str(item['id']) in selected]
        results = []

        with jobs_lock:
            jobs[job_id].update({"status": "running", "total": len(items)})

        for position, item in enumerate(items, start=1):
            with jobs_lock:
                jobs[job_id].update({"current": position, "current_id": item['id']})
            rel_path = get_server_relative_path(main_path_url, item['path_sub'])
            try:
                res = process_sharepoint_permission(
                    site_url=site_url,
                    relative_folder_path=rel_path,
                    email=item['email'],
                    action_type=action,
                    role_name=item['role'],
                    start_date=item.get('start_date', ''),
                    end_date=item.get('end_date', '')
                )
            except Exception as error:
                res = {"status": "Error", "remark": f"{type(error).__name__}: {error}"}
            result = {
                "id": item['id'],
                "status": res['status'],
                "email_status": res.get('email_status', 'not_sent'),
                "logs_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "remark": res['remark']
            }
            results.append(result)
            with jobs_lock:
                jobs[job_id]["results"] = results.copy()
            if position < len(items):
                time.sleep(grant_delay_seconds)

        update_excel_logs(results)
        with jobs_lock:
            jobs[job_id].update({"status": "completed", "results": results})
    except Exception as error:
        with jobs_lock:
            jobs[job_id].update({"status": "failed", "error": f"{type(error).__name__}: {error}"})

@app.route('/api/execute/<job_id>', methods=['GET'])
def execute_status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"success": False, "error": "ไม่พบงานที่กำลังประมวลผล"}), 404
        return jsonify({"success": True, **job})

@app.route('/api/permissions/<int:permission_id>', methods=['PUT'])
def edit_permission(permission_id):
    try:
        update_permission(permission_id, request.json or {})
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route('/api/permissions/<int:permission_id>/copy', methods=['POST'])
def duplicate_permission(permission_id):
    try:
        new_id = copy_permission(permission_id)
        return jsonify({"success": True, "id": new_id})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route('/api/permissions/<int:permission_id>', methods=['DELETE'])
def remove_permission(permission_id):
    try:
        delete_permission(permission_id)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

if __name__ == '__main__':
    app.run(port=5000, debug=True)