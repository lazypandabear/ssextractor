from flask import Flask, render_template, request
import threading  # Optional: for running migration in background
import main  # Import your main.py module

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        config = {
            "smartsheet_api_key": request.form.get('smartsheet_api_key'),
            "smartsheet_folder_id": request.form.get('smartsheet_folder_id'),
            "google_drive_sheets_folder_id": request.form.get('google_drive_sheets_folder_id'),
            "google_drive_comments_folder_id": request.form.get('google_drive_comments_folder_id'),
            "google_drive_attachments_folder_id": request.form.get('google_drive_attachments_folder_id')
        }
        
        # Option 1: Run synchronously (the user waits until the process completes)
        result = main.run_migration(config)
        
        # Option 2 (Recommended for responsiveness): Run in a background thread
        # def background_migration():
        #     main.run_migration(config)
        # threading.Thread(target=background_migration).start()
        # result = "Migration has started in the background."
        
        return render_template('result.html', result=result, config=config)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
