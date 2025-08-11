import os
import shutil
import zipfile
import json
import re
import html
from datetime import date, datetime
from flask import Flask, render_template, request, send_file, jsonify, Response
from email.message import EmailMessage

app = Flask(__name__, static_folder="static", template_folder="templates")

BASE_DIR      = os.path.dirname(__file__)
TEMPLATE_HTML = os.path.join(BASE_DIR, "template.html")  # Using your template.html file
TEMP_ROOT     = os.path.join(BASE_DIR, "rollup_temp")
SETTINGS_FILE = os.path.join(BASE_DIR, "form_data.json")

def sanitize(name: str) -> str:
    for c in r'\/:*?"<>|':
        name = name.replace(c, "_")
    return name.strip()

def create_eml_file(to_addresses, cc_addresses, subject, html_body, filename):
    """Create only a .eml draft file - no other files"""
    
    from email.message import EmailMessage
    import re
    
    msg = EmailMessage()
    
    # Critical header that makes it a draft
    msg["X-Unsent"] = "1"
    
    # Set subject
    msg["Subject"] = subject
    
    # Set recipients only if they exist
    if to_addresses.strip():
        msg["To"] = to_addresses
    if cc_addresses and cc_addresses.strip():
        msg["Cc"] = cc_addresses
    
    # Create clean plain text version
    plain_body = html_body
    # Remove HTML tags
    plain_body = re.sub(r'<[^>]+>', '', plain_body)
    # Clean up HTML entities
    plain_body = plain_body.replace('&nbsp;', ' ')
    plain_body = plain_body.replace('&amp;', '&')
    plain_body = plain_body.replace('&lt;', '<')
    plain_body = plain_body.replace('&gt;', '>')
    plain_body = plain_body.replace('&quot;', '"')
    # Remove any = encoding artifacts
    plain_body = re.sub(r'=[A-Fa-f0-9]{2}', '', plain_body)
    plain_body = re.sub(r'=\w+', '', plain_body)  # Remove =sdf, =dfg etc
    # Clean up whitespace
    plain_body = re.sub(r'\n\s*\n', '\n\n', plain_body)
    plain_body = plain_body.strip()
    
    # Clean HTML body too
    clean_html = html_body
    clean_html = re.sub(r'=[A-Fa-f0-9]{2}', '', clean_html)
    clean_html = re.sub(r'=\w+', '', clean_html)
    
    # Set content with 8bit transfer encoding to avoid quoted-printable
    msg.set_content(plain_body, charset='utf-8', cte='8bit')
    msg.add_alternative(clean_html, subtype="html", charset='utf-8', cte='8bit')
    
    # Write as bytes
    with open(filename, 'wb') as f:
        f.write(bytes(msg))
    
    # Return only the .eml filename (not a list)
    return filename

@app.route("/", methods=["GET"])
def form():
    return render_template("form.html")

@app.route("/load-settings", methods=["GET"])
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = {}
    return jsonify(data)

@app.route("/save-settings", methods=["POST"])
def save_settings():
    data = request.get_json(force=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    now = datetime.now().isoformat(sep=' ', timespec='seconds')
    print(f"[{now}] Saved settings, keys: {list(data.keys())}")
    return jsonify({"status":"ok", "saved_at": now}), 200

@app.route("/backup-settings", methods=["GET"])
def backup_settings():
    """Export all settings as JSON for backup"""
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        data = {}
    
    # Add timestamp to backup
    data["backup_timestamp"] = datetime.now().isoformat()
    data["backup_version"] = "1.0"
    
    # Create JSON response for download
    json_data = json.dumps(data, indent=2, ensure_ascii=False)
    
    return Response(
        json_data,
        mimetype='application/json',
        headers={
            'Content-Disposition': f'attachment; filename=rollup_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
        }
    )

@app.route("/restore-settings", methods=["POST"])
def restore_settings():
    """Restore settings from uploaded JSON backup"""
    try:
        if 'backup_file' not in request.files:
            return jsonify({"status": "error", "message": "No file uploaded"}), 400
        
        file = request.files['backup_file']
        if file.filename == '':
            return jsonify({"status": "error", "message": "No file selected"}), 400
        
        if not file.filename.endswith('.json'):
            return jsonify({"status": "error", "message": "Please upload a JSON file"}), 400
        
        # Read and parse JSON
        content = file.read().decode('utf-8')
        data = json.loads(content)
        
        # Remove backup metadata before saving
        data.pop("backup_timestamp", None)
        data.pop("backup_version", None)
        
        # Save restored settings
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        
        now = datetime.now().isoformat(sep=' ', timespec='seconds')
        print(f"[{now}] Restored settings from backup, keys: {list(data.keys())}")
        
        return jsonify({
            "status": "success", 
            "message": "Settings restored successfully",
            "restored_at": now
        }), 200
        
    except json.JSONDecodeError:
        return jsonify({"status": "error", "message": "Invalid JSON file"}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": f"Error restoring backup: {str(e)}"}), 500

@app.route("/generate", methods=["POST"])
def generate_zip():
    # 1) Rollup date
    rd = request.form.get("Date","").strip()
    if not rd:
        t = date.today()
        rd = f"{t.year}-{t.month:02d}-{t.day:02d}"
    yyyy, mm, dd = rd.split("-")
    date_fmt = f"{mm}-{dd}-{yyyy}"
    add_nosend = (request.form.get("GlobalNosend") == "true")

    # 2) Which rows
    rows = sorted({
        int(k.rsplit("_",1)[1])
        for k in request.form
        if k.startswith("CustomerName_") and k.rsplit("_",1)[1].isdigit()
    })

    # 3) Reset temp folder
    if os.path.exists(TEMP_ROOT):
        shutil.rmtree(TEMP_ROOT)
    os.makedirs(TEMP_ROOT, exist_ok=True)

    # 4) Load HTML template instead of exporting from Word
    with open(TEMPLATE_HTML, "r", encoding="utf-8") as f:
        master_html = f.read()
    
    css = """
      <style>
        body, p, ul, li { font-size:11pt !important; line-height:1.15 !important; }
      </style>
    """
    master_html = css + master_html

    # 5) Generate emails (same logic as before)
    generated = []

    for i in rows:
        cust = request.form.get(f"CustomerName_{i}","").strip()
        if not cust:
            continue
        safe = sanitize(cust)

        # simple placeholders
        simple = {
            "{{CustomerName}}":       html.escape(cust),
            "{{Date}}":               html.escape(date_fmt),
            "{{Contacts}}":           html.escape(request.form.get(f"ASContacts_{i}","").strip()),
            "{{CCAddresses}}":        html.escape(request.form.get(f"AccountTeamContacts_{i}","").strip()),
            "{{AdditionalContacts}}": html.escape(request.form.get(f"AdditionalContacts_{i}","").strip())
        }

        # Get section titles (editable) 
        general_title = request.form.get(f"GeneralTitle_{i}","General").strip()
        account_title = request.form.get(f"AccountMgmtTitle_{i}","Account Management").strip()
        engineer_title = request.form.get(f"DesignatedEngTitle_{i}","Designated Engineer(s)").strip()
        notes_title = request.form.get(f"SpecialNotesTitle_{i}","Special Notes").strip()
        
        # List or rich-html placeholders (same as before)
        sections = {
            "{{DiscussionTopics}}": request.form.get(f"DiscussionTopics_{i}",""),
            "{{AccountManagement}}": request.form.get(f"AccountManagement_{i}",""),
            "{{DesignatedEngineer}}": request.form.get(f"DesignatedEngineer_{i}",""),
            "{{SpecialNotes}}": request.form.get(f"SpecialNotes_{i}",""),
        }

        # Handle custom sections dynamically (including section count)
        custom_sections = []
        section_count = 0
        j = 1
        while j <= 20:  # Limit to reasonable number
            custom_title = request.form.get(f"CustomSection{j}Title_{i}","").strip()
            custom_content = request.form.get(f"CustomSection{j}_{i}","").strip()
            if custom_title and custom_content:
                custom_sections.append((custom_title, custom_content))
                section_count = max(section_count, j)
            j += 1

        # get raw signature and clean it
        raw_sig = request.form.get(f"Signature_{i}","").strip()
        
        # Clean up encoding artifacts in signature too
        raw_sig = re.sub(r'=[A-Fa-f0-9]{2}', '', raw_sig)
        raw_sig = re.sub(r'=\w+', '', raw_sig)
        raw_sig = raw_sig.replace('+', ' ')
        
        if raw_sig.startswith("<"):
            # style each <p> to remove spacing
            signature_html = re.sub(
                r'<p>',
                '<p style="margin:0; line-height:1.0;">',
                raw_sig
            )
        else:
            signature_html = (
                f'<p style="margin:0; line-height:1.0;">{html.escape(raw_sig)}</p>'
                if raw_sig else ""
            )

        # Build HTML - replace placeholders the simple way
        html_body = master_html
        
        # Replace simple placeholders
        for ph, val in simple.items():
            html_body = html_body.replace(ph, val)
        
        # Replace section titles first
        html_body = html_body.replace(">General<", f">{html.escape(general_title)}<")
        html_body = html_body.replace(">Account Management<", f">{html.escape(account_title)}<")
        html_body = html_body.replace(">Designated Engineer(s)<", f">{html.escape(engineer_title)}<")
        html_body = html_body.replace(">Special Notes<", f">{html.escape(notes_title)}<")
        
        # Replace section content
        for ph, raw in sections.items():
            content = raw.strip()
            
            # Clean up any encoding artifacts first
            content = re.sub(r'=[A-Fa-f0-9]{2}', '', content)  # Remove =XX patterns
            content = re.sub(r'=\w+', '', content)  # Remove =sdf, =dfg etc
            content = content.replace('+', ' ')  # Replace + with spaces
            
            if content.startswith("<"):
                replacement = content
            else:
                lines = [l.strip() for l in content.splitlines() if l.strip()]
                if lines:
                    items = "".join(f"<li>{html.escape(l)}</li>" for l in lines)
                    replacement = f"<ul>{items}</ul>"
                else:
                    replacement = ""
            html_body = html_body.replace(ph, replacement)

        # Add signature and custom sections at the end
        if signature_html or custom_sections:
            additions = ""
            
            # Add custom sections to the discussion table first
            if custom_sections:
                custom_sections_html = ""
                for title, content in custom_sections:
                    # Clean up encoding artifacts in custom content
                    content = re.sub(r'=[A-Fa-f0-9]{2}', '', content)
                    content = re.sub(r'=\w+', '', content)
                    content = content.replace('+', ' ')
                    
                    if content.startswith("<"):
                        section_content = content
                    else:
                        lines = [l.strip() for l in content.splitlines() if l.strip()]
                        if lines:
                            items = "".join(f"<li>{html.escape(l)}</li>" for l in lines)
                            section_content = f"<ul>{items}</ul>"
                        else:
                            section_content = ""
                    
                    if section_content:
                        custom_sections_html += f"""
                        <tr>
                          <td style="background:#c0392b;color:#ffffff;font-weight:bold;text-align:center;padding:3px 6px;border-bottom:1px solid #8b2b21;border-left:1px solid #8b2b21;border-right:1px solid #8b2b21;font-size:11pt;">{html.escape(title)}</td>
                        </tr>
                        <tr>
                          <td style="padding:6px;border-bottom:1px solid #8b2b21;border-left:1px solid #8b2b21;border-right:1px solid #8b2b21;font-size:11pt;">{section_content}</td>
                        </tr>"""
                
                # Add custom sections immediately after Special Notes (no gap)
                if custom_sections_html:
                    # Find the Special Notes section specifically and add right after it
                    special_notes_end = html_body.find('</tr>', html_body.find('Special Notes'))
                    if special_notes_end != -1:
                        # Find the next </tr> after Special Notes content
                        content_end = html_body.find('</tr>', special_notes_end + 5)
                        if content_end != -1:
                            # Insert custom sections right after Special Notes content
                            html_body = html_body[:content_end + 5] + custom_sections_html + html_body[content_end + 5:]
                    else:
                        # Fallback: add before the last </table>
                        last_table_end = html_body.rfind("</table>")
                        if last_table_end != -1:
                            html_body = html_body[:last_table_end] + custom_sections_html + html_body[last_table_end:]
            
            # Add signature at the very end
            if signature_html:
                if "</body>" in html_body:
                    signature_section = f'<div style="margin-top: 30px;">{signature_html}</div>'
                    html_body = html_body.replace("</body>", f"{signature_section}</body>")
                else:
                    signature_section = f'<div style="margin-top: 30px;">{signature_html}</div>'
                    html_body += signature_section

        # build CC
        cc = simple["{{CCAddresses}}"]
        addl = simple["{{AdditionalContacts}}"]
        if addl:
            cc = f"{cc}; {addl}" if cc else addl
        if add_nosend:
            cc = f"{cc}; nosend" if cc else "nosend"

        # create only .eml file (no .oft or .txt files)
        subject = f"{cust} Weekly Rollup {date_fmt}"
        eml_path = os.path.join(TEMP_ROOT, f"{safe}_{date_fmt}.eml")
        
        create_eml_file(
            to_addresses=simple["{{Contacts}}"].replace("&amp;", "&"),
            cc_addresses=cc.replace("&amp;", "&") if cc else "",
            subject=subject,
            html_body=html_body,
            filename=eml_path
        )
        
        generated.append((os.path.basename(eml_path), eml_path))

    # 6) Zip with date in filename
    zip_filename = f"Rollup_Messages_{date_fmt}.zip"
    zip_path = os.path.join(TEMP_ROOT, zip_filename)
    with zipfile.ZipFile(zip_path, "w") as zf:
        for name, path in generated:
            zf.write(path, arcname=name)

    return send_file(
        zip_path,
        as_attachment=True,
        download_name=zip_filename,
        mimetype="application/zip"
    )

if __name__ == "__main__":
    app.run(debug=False, port=5000)