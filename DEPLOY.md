## How to Reset & Upload Correctly (Fresh Start)
If your deployment is failing, it's best to start fresh to ensure the folder structure is perfect.

### Step 1: Delete & Re-create Repository
1.  Go to your GitHub repository **Settings** (top right tab of the repo).
2.  Scroll to the very bottom to the **Danger Zone**.
3.  Click **Delete this repository**.
4.  Type the repo name to confirm and delete it.
5.  Go to [github.com/new](https://github.com/new) and create the repository again (same name is fine).

### Step 2: The "Clean" Upload
1.  In your new repository, click **"uploading an existing file"**.
2.  Open your project folder on your computer.

#### ✅ DRAG THESE IN (Drag the FOLDERS themselves):
*   `templates` (The entire folder. Do not open it.)
*   `static` (The entire folder. Do not open it.)
*   `app.py`
*   `reconciliation_v2.py`
*   `requirements.txt`
*   `Procfile`

#### ❌ DO NOT UPLOAD (Skip these):
*   `uploads` folder
*   `results` folder
*   `__pycache__` folder
*   `.env` or `venv` folders
*   Any Excel files (`.xlsx`)

### Step 3: Redeploy
1.  Go to your Render Dashboard.
2.  It might say "Deploy Failed". Click on your service.
3.  Go to **Settings** and scroll down to **Build & Deploy**.
4.  Click **Clear Build Cache & Deploy**.

## Step 3: Wait for Build
Render will now build your application. It might take a few minutes. Once done, you will see a URL like `https://gst-reco.onrender.com`.

## Note on File Storage
Since this is a free hosting service, the file system is **ephemeral**.
*   Files uploaded and results generated will stay for a short while but will be deleted if the app restarts (which happens automatically when idle).
*   This is fine for a tool where you "Upload > Process > Download" immediately.

## Troubleshooting
If the build fails, check the "Logs" tab in Render for error messages.
