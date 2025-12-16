from flask import Flask, render_template, request, jsonify
import threading
import main
import process_state
import config

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Reset cancel flag for each new migration
        process_state.cancel_requested = False
        
        # Get configuration from form
        configuration = {
            "SMARTSHEET_API_KEY": request.form.get('smartsheet_api_key'),
            "SMARTSHEET_FOLDER_ID": request.form.get('smartsheet_folder_id'),
            "GOOGLE_DRIVE_SHEETS_FOLDER_ID": request.form.get('google_drive_sheets_folder_id'),
            "GOOGLE_DRIVE__COMMENTS_FOLDER_ID": request.form.get('google_drive_comments_folder_id'),
            "GOOGLE_DRIVE_ATTACHMENTS_FOLDER_ID": request.form.get('google_drive_attachments_folder_id')
        }

        # Update global configuration
        config.CREDENTIALS.update(configuration)
        #print(config.CREDENTIALS)
        # Start migration in a background thread
        threading.Thread(target=main.run_migration).start()
        return render_template('migration_started.html')
    return render_template('index.html')

@app.route('/status', methods=['GET'])
def status():
    # Return current migration status as JSON
    return jsonify(process_state.migration_status)

@app.route('/cancel', methods=['POST'])
def cancel():
    process_state.cancel_requested = True
    return jsonify({"status": "cancelled"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
