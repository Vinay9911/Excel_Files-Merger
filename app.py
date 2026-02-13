from flask import Flask, render_template, request, Response, send_file, jsonify
import xlwings as xw
import os
import uuid
import shutil
import json

app = Flask(__name__)

JOB_DIR = ".merger_jobs"
os.makedirs(JOB_DIR, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    uploaded_files = request.files.getlist("files")
    if not uploaded_files or uploaded_files[0].filename == '':
        return jsonify({"error": "No files selected."}), 400

    job_id = str(uuid.uuid4())
    job_path = os.path.join(JOB_DIR, job_id)
    os.makedirs(job_path, exist_ok=True)

    saved_files = []
    for file in uploaded_files:
        file_path = os.path.join(job_path, file.filename)
        file.save(file_path)
        saved_files.append(file.filename)

    return jsonify({"job_id": job_id, "files": saved_files})

@app.route('/stream/<job_id>')
def stream(job_id):
    job_path = os.path.join(JOB_DIR, job_id)
    custom_filename = request.args.get('filename', 'Merged_Workbooks')
    if not custom_filename.endswith('.xlsx'):
        custom_filename += '.xlsx'
        
    merged_path = os.path.join(job_path, custom_filename)

    def generate():
        excel_app = xw.App(visible=False)
        excel_app.display_alerts = False
        error_log = []
        total_sheets = 0
        files_processed = 0

        try:
            yield f"data: {json.dumps({'log': 'Initializing Excel Engine...', 'type': 'system'})}\n\n"
            merged_wb = excel_app.books.add()
            
            yield f"data: {json.dumps({'log': 'Creating Master Index Table of Contents...', 'type': 'system'})}\n\n"
            toc_sheet = merged_wb.sheets.add('Table of Contents', before=merged_wb.sheets[0])
            toc_sheet.range('A1').value = "Master Index"
            toc_sheet.range('A1').font.bold = True
            toc_sheet.range('A1').font.size = 18
            toc_row = 3

            files = sorted(os.listdir(job_path))
            for filename in files:
                if filename == custom_filename: continue 
                
                path = os.path.join(job_path, filename)
                
                # Assigning string to a variable first prevents the f-string backslash error
                msg_open = f"Opening File: {filename}"
                yield f"data: {json.dumps({'log': msg_open, 'type': 'file'})}\n\n"
                
                try:
                    source_wb = excel_app.books.open(path)
                    files_processed += 1
                except Exception:
                    err = f"Failed to open '{filename}'"
                    error_log.append(err)
                    yield f"data: {json.dumps({'log': err, 'type': 'error'})}\n\n"
                    continue

                for sheet in source_wb.sheets:
                    try:
                        try: rows, cols = sheet.used_range.shape
                        except: rows, cols = ("?", "?")

                        msg_sheet = f"  ↳ Copying: {sheet.name} (Verified: {rows} Rows x {cols} Cols)"
                        yield f"data: {json.dumps({'log': msg_sheet, 'type': 'sheet'})}\n\n"
                        
                        sheet.api.Copy(After=merged_wb.sheets[-1].api)
                        total_sheets += 1
                        
                        toc_sheet.api.Hyperlinks.Add(
                            Anchor=toc_sheet.range(f'A{toc_row}').api, 
                            Address="", 
                            SubAddress=f"'{sheet.name}'!A1", 
                            TextToDisplay=f"{filename} - {sheet.name}"
                        )
                        toc_row += 1

                    except Exception as e:
                        err = f"Failed to copy '{sheet.name}'"
                        error_log.append(err)
                        yield f"data: {json.dumps({'log': err, 'type': 'error'})}\n\n"
                
                source_wb.close()

            if len(merged_wb.sheets) > 1 and merged_wb.sheets[1].name == 'Sheet1':
                merged_wb.sheets[1].delete()
            
            merged_wb.save(merged_path)
            merged_wb.close()

            # Clean formatting for the end summary
            yield f"data: {json.dumps({'log': '====================================', 'type': 'summary'})}\n\n"
            
            if not error_log:
                yield f"data: {json.dumps({'log': '✅ MERGE SUCCESSFULLY COMPLETED!', 'type': 'success'})}\n\n"
            else:
                yield f"data: {json.dumps({'log': '⚠️ MERGE COMPLETED WITH WARNINGS.', 'type': 'warning'})}\n\n"
            
            yield f"data: {json.dumps({'log': f'Files Processed: {files_processed}', 'type': 'summary'})}\n\n"
            yield f"data: {json.dumps({'log': f'Total Sheets Combined: {total_sheets}', 'type': 'summary'})}\n\n"
            yield f"data: {json.dumps({'log': '====================================', 'type': 'summary'})}\n\n"

            yield f"data: {json.dumps({'status': 'complete', 'filename': custom_filename})}\n\n"

        except Exception as e:
            err_msg = f"FATAL ERROR: {str(e)}"
            yield f"data: {json.dumps({'log': err_msg, 'type': 'error'})}\n\n"
            yield f"data: {json.dumps({'status': 'error'})}\n\n"
        finally:
            excel_app.quit()

    return Response(generate(), mimetype='text/event-stream')

@app.route('/download/<job_id>')
def download(job_id):
    filename = request.args.get('filename')
    job_path = os.path.join(JOB_DIR, job_id)
    file_path = os.path.join(job_path, filename)

    response = send_file(
        file_path,
        download_name=filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    @response.call_on_close
    def cleanup():
        shutil.rmtree(job_path, ignore_errors=True)
        
    return response

if __name__ == '__main__':
    app.run(debug=True)