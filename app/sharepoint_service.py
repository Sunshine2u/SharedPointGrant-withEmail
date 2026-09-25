import requests
import urllib.parse
import os
from datetime import datetime
from html import escape
from dotenv import load_dotenv

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
load_dotenv(os.path.join(PROJECT_ROOT, "config", ".env"))

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "").strip()
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")
SENDER_EMAIL = os.getenv("GROUP_SENDER_EMAIL", "").strip()
THAI_MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]

def format_thai_short_date(value):
    try:
        parsed = datetime.strptime(str(value)[:10], "%Y-%m-%d")
        return f"{parsed.day} {THAI_MONTHS[parsed.month - 1]} {str(parsed.year)[-2:]}"
    except (TypeError, ValueError):
        return str(value or "")

def get_sharepoint_cookies():
    rtFa = os.getenv("SP_RTFA", "").strip()
    fed_auth = os.getenv("SP_FEDAUTH", "").strip()
    if rtFa and fed_auth:
        return {"rtFa": rtFa, "FedAuth": fed_auth}

    # ป้อน Cookie แบบ Manual หากยังไม่ได้ตั้งค่าใน .env
    print("\n" + "="*60)
    print(" >>> กรุณาคัดลอกค่า Cookie จาก Browser (Edge/Chrome) มาวาง")
    print("     (เปิด SharePoint ในเว็บ -> F12 -> Application -> Cookies)")
    print("="*60)
    
    rtFa = input("กรอกค่า rtFa: ").strip()
    FedAuth = input("กรอกค่า FedAuth: ").strip()

    return {
        "rtFa": rtFa,
        "FedAuth": FedAuth
    }

def get_server_relative_path(main_path_url, sub_path):
    parsed_main = urllib.parse.urlparse(main_path_url)
    path_decoded = urllib.parse.unquote(parsed_main.path)
    
    if "/sites/" in path_decoded:
        base_path = "/sites/" + path_decoded.split("/sites/")[1]
    else:
        base_path = path_decoded

    clean_sub = sub_path.replace("\\", "/").strip("/")
    full_path = f"{base_path}/{clean_sub}".rstrip("/")
    return full_path

def build_sharepoint_folder_url(site_url, folder_path):
    parsed = urllib.parse.urlparse(site_url)
    domain_url = f"{parsed.scheme}://{parsed.netloc}"
    encoded_path = urllib.parse.quote(folder_path.lstrip("/"), safe="/")
    return f"{domain_url}/{encoded_path}"

def get_graph_access_token():
    if not AZURE_TENANT_ID or not AZURE_CLIENT_ID or not AZURE_CLIENT_SECRET:
        raise ValueError("ยังไม่ได้ตั้งค่า AZURE_TENANT_ID, AZURE_CLIENT_ID หรือ AZURE_CLIENT_SECRET ในไฟล์ .env")

    token_url = f"https://login.microsoftonline.com/{urllib.parse.quote(AZURE_TENANT_ID, safe='')}/oauth2/v2.0/token"
    response = requests.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": AZURE_CLIENT_ID,
            "client_secret": AZURE_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=15,
    )
    if response.status_code != 200:
        try:
            error_data = response.json()
            error_detail = error_data.get("error_description") or error_data.get("error")
        except ValueError:
            error_detail = None
        detail = f": {error_detail}" if error_detail else ""
        raise RuntimeError(
            f"ขอ Microsoft Graph access token ไม่สำเร็จ (HTTP {response.status_code}){detail}"
        )

    access_token = response.json().get("access_token")
    if not access_token:
        raise RuntimeError("Microsoft Graph ไม่ได้ส่ง access token กลับมา")
    return access_token


def check_graph_connection():
    try:
        if not SENDER_EMAIL:
            return False, "ยังไม่ได้ตั้งค่า GROUP_SENDER_EMAIL ในไฟล์ .env"

        get_graph_access_token()
        return True, f"เชื่อมต่อ Microsoft Graph สำเร็จ: {SENDER_EMAIL}"
    except Exception as e:
        return False, f"เชื่อมต่อ Microsoft Graph ไม่ได้ ({type(e).__name__}): {e}"

def process_sharepoint_permission(site_url, relative_folder_path, email, action_type, role_name, start_date="", end_date=""):
    graph_ready, graph_message = check_graph_connection()
    print(f"[Microsoft Graph] {graph_message}")
    if not graph_ready:
        return {
            "status": "Error",
            "email_status": "not_checked",
            "remark": graph_message
        }

    cookies = get_sharepoint_cookies()
    
    # ดึง Request Digest Token
    context_url = f"{site_url}/_api/contextinfo"
    headers_ctx = {"Accept": "application/json;odata=verbose"}
    
    try:
        res_ctx = requests.post(context_url, cookies=cookies, headers=headers_ctx)
        if res_ctx.status_code != 200:
            # ลบบรรทัด os.remove(COOKIE_FILE) ออกแล้ว
            return {"status": "Error", "remark": f"Cookie หมดอายุ หรือ site_url ผิด (HTTP {res_ctx.status_code})"}
            
        form_digest = res_ctx.json()['d']['GetContextWebInformation']['FormDigestValue']
    except Exception as e:
        # ลบบรรทัด os.remove(COOKIE_FILE) ออกแล้ว
        return {"status": "Error", "remark": f"Auth Failed: {str(e)}"}

    headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": form_digest
    }

    encoded_path = urllib.parse.quote(relative_folder_path)
    folder_url = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{encoded_path}')/ListItemAllFields"

    # Break Inheritance
    break_url = f"{folder_url}/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)"
    requests.post(break_url, cookies=cookies, headers=headers)

    # Get Principal ID จาก Email
    ensure_url = f"{site_url}/_api/web/ensureuser"
    user_res = requests.post(ensure_url, cookies=cookies, headers=headers, json={"logonName": email})
    if user_res.status_code not in [200, 201]:
        return {"status": "Error", "remark": f"User not found: {email}"}
    
    principal_id = user_res.json()['d']['Id']

    if action_type == "GRANT":
        role_def_id = 1073741827 if "Edit" in role_name else 1073741826
        add_url = f"{folder_url}/roleassignments/addroleassignment(principalid={principal_id}, roledefid={role_def_id})"
        res = requests.post(add_url, cookies=cookies, headers=headers)
        
        if res.status_code in [200, 201, 204]:
            status_text = "Edit" if role_def_id == 1073741827 else "Read Only"
            folder_url = build_sharepoint_folder_url(site_url, relative_folder_path)
            email_result = send_graph_email(email, relative_folder_path, status_text, "GRANT", folder_url, start_date, end_date)
            return {
                "status": status_text,
                "email_status": email_result["status"],
                "remark": f"Granted successfully; {email_result['remark']}"
            }
        return {"status": "Error", "remark": f"GRANT Failed (Code: {res.status_code})"}
            
    elif action_type == "REVOKE":
        remove_url = f"{folder_url}/roleassignments/getbyprincipalid({principal_id})"
        delete_headers = headers.copy()
        delete_headers["IF-MATCH"] = "*"
        delete_headers["X-HTTP-Method"] = "DELETE"
        
        res = requests.post(remove_url, cookies=cookies, headers=delete_headers)
        
        # 200, 204 คือลบสำเร็จ | 404 คือไม่มีสิทธิ์เดิมอยู่แล้ว
        if res.status_code in [200, 204, 404, 500]:
            email_result = send_graph_email(email, relative_folder_path, "None", "REVOKE")
            return {
                "status": "ยกเลิก",
                "email_status": email_result["status"],
                "remark": f"Revoked successfully; {email_result['remark']}"
            }

        return {"status": "Error", "remark": f"REVOKE Failed (Code: {res.status_code})"}

    return {"status": "Error", "remark": "Unknown Action"}


def send_graph_email(email_to, folder_name, role_name, action_type, folder_url=None, start_date="", end_date=""):
    try:
        folder_subject = folder_name.rstrip("/").split("/")[-1]
        if action_type == "GRANT":
            status_title = "Access Granted"
            action_desc = "Access Granted"
            subject = f"[Budget Planning] Notification - Access Granted : {folder_subject}"
        else:
            status_title = "Access Revoked"
            action_desc = "Access Revoked"
            subject = f"[Budget Planning] Notification - Access Revoked : {folder_subject}"

        folder_display_name = escape(folder_name.rstrip("/").split("/")[-1])
        role_colors = {
            "Edit": ("#166534", "#dcfce7"),
            "Read Only": ("#0369a1", "#e0f2fe"),
            "None": ("#b91c1c", "#fee2e2"),
        }
        role_text = str(role_name or "")
        role_foreground, role_background = role_colors.get(role_text, ("#374151", "#f3f4f6"))
        role_display_html = (
            f'<strong style="color: {role_foreground}; background-color: {role_background}; '
            f'padding: 4px 10px; border-radius: 4px; display: inline-block;">'
            f'{escape(role_text)}</strong>'
        )
        folder_link_html = ""
        resource_name_html = folder_display_name
        period_html = ""
        if action_type == "GRANT" and folder_url:
            safe_folder_url = escape(folder_url, quote=True)
            resource_name_html = f'<a href="{safe_folder_url}" target="_blank" style="color: #0078d4; font-weight: bold; text-decoration: underline;">{folder_display_name}</a>'
            folder_link_html = f"""
                    <tr>
                        <td colspan="2" style="padding: 15px 8px 5px 8px;">
                            <a href="{safe_folder_url}" target="_blank" style="background-color: #0078d4; color: white; padding: 10px 18px; text-decoration: none; border-radius: 4px; display: inline-block; font-weight: bold;">Open Folder in SharePoint</a>
                        </td>
                    </tr>
            """

            period_html = f"""
                    <tr><td style="padding: 8px;"><b>ช่วงเวลาให้สิทธิ์:</b></td><td style="padding: 8px;">{escape(format_thai_short_date(start_date))} ถึง {escape(format_thai_short_date(end_date))}</td></tr>
            """

        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
                <div style="background-color: #0078d4; color: white; padding: 20px; font-size: 18px; font-weight: bold;">
                Budget Planning - Notification
            </div>
            <div style="padding: 20px; border: 1px solid #ddd;">
                <p>Dear User,</p>
                <h3>{status_title}</h3>
                <p>แจ้งสถานะการเข้าถึงโฟลเดอร์ <b>Budget Planning</b>.</p>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 8px; width: 30%;"><b>Action:</b></td><td style="padding: 8px;">{action_desc}</td></tr>
                    <tr><td style="padding: 8px;"><b>ชื่อโฟลเดอร์:</b></td><td style="padding: 8px;">{resource_name_html}</td></tr>
                    <tr><td style="padding: 8px;"><b>สิทธิ์การเข้าถึง:</b></td><td style="padding: 8px;">{role_display_html}</td></tr>
                    {period_html}
                    {folder_link_html}
                </table>
                <br>
                <p style="color: #666; font-size: 12px;">This is an automated notification from Budget Planning.</p>
            </div>
        </body>
        </html>
        """
        
        token = get_graph_access_token()
        sender_path = urllib.parse.quote(SENDER_EMAIL, safe="")
        endpoint = f"https://graph.microsoft.com/v1.0/users/{sender_path}/sendMail"
        response = requests.post(
            endpoint,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json={
                "message": {
                    "subject": subject,
                    "body": {"contentType": "HTML", "content": html_body},
                    "toRecipients": [{"emailAddress": {"address": email_to}}],
                },
                "saveToSentItems": True,
            },
            timeout=15,
        )
        if response.status_code != 202:
            raise RuntimeError(f"Microsoft Graph sendMail ล้มเหลว (HTTP {response.status_code})")

        print(f"[Email Success] ส่งอีเมลแจ้งเตือนไปยัง {email_to} เรียบร้อยแล้ว")
        return {"status": "sent", "remark": f"ส่งอีเมลสำเร็จไปยัง {email_to}"}
    except Exception as e:
        error_message = f"ส่งอีเมลไม่สำเร็จ ({type(e).__name__}): {e}"
        print(f"[Email Failed] {error_message}")
        return {"status": "failed", "remark": error_message}
