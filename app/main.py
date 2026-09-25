from flask import Flask, render_template, jsonify, request
from excel_service import read_excel_data, update_excel_logs, update_permission, copy_permission, delete_permission
from sharepoint_service import process_sharepoint_permission, get_server_relative_path
from datetime import datetime
import urllib.parse


app = Flask(__name__)

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

        excel_info = read_excel_data()
        main_path_url = excel_info['main_path']

        # แทนที่บรรทัด site_url เดิมใน main.py ด้วยโค้ดนี้
        parsed_main = urllib.parse.urlparse(main_path_url)
        site_name = parsed_main.path.split('/sites/')[1].split('/')[0] if '/sites/' in parsed_main.path else ''
        site_url = f"{parsed_main.scheme}://{parsed_main.netloc}/sites/{site_name}"

        results = []
        current_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for item in excel_info['permissions']:
            # เทียบ ID โดยแปลงเป็น int/str ให้ตรงกัน
            if str(item['id']) in [str(x) for x in selected_ids]:
                rel_path = get_server_relative_path(main_path_url, item['path_sub'])
                
                res = process_sharepoint_permission(
                    site_url=site_url,
                    relative_folder_path=rel_path,
                    email=item['email'],
                    action_type=action,
                    role_name=item['role'],
                    start_date=item.get('start_date', ''),
                    end_date=item.get('end_date', '')
                )
                
                results.append({
                    "id": item['id'],
                    "status": res['status'],
                    "email_status": res.get('email_status', 'not_sent'),
                    "logs_date": current_now,
                    "remark": res['remark']
                })

        # อัปเดตลง Excel
        update_excel_logs(results)
        return jsonify({"success": True, "results": results})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

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