# n8n Gmail PDF Workflow

This project runs a Dockerized n8n instance plus a small Node.js controller. The controller creates an n8n workflow, triggers it, fetches the latest PDF attachment from Gmail, and saves it into this project's `downloads/` folder.

## What Runs

- `n8n`: official n8n Docker image, exposed at `http://localhost:5678`
- `my_project`: this Node.js controller, built from the local `Dockerfile`
- `downloads/`: shared folder mounted into n8n at `/home/node/downloads`
- `n8n_data`: Docker volume that stores n8n credentials, API keys, and workflow data

The setup uses relative paths, so it works from the same project folder on macOS, Windows, and Linux with Docker Desktop or Docker Engine.

## One-Time Setup

Create your env file:

```bash
cp .env.example .env
```

On Windows PowerShell:

```powershell
Copy-Item .env.example .env
```

Create the downloads folder if it does not already exist:

```bash
mkdir -p downloads
```

On Windows PowerShell:

```powershell
New-Item -ItemType Directory -Force downloads
```

Start n8n first:

```bash
docker compose up --build n8n
```

Open `http://localhost:5678`, then:

1. Create or sign in to the local n8n owner account.
2. Go to Settings, then n8n API or API Keys, and create an API key.
3. Create a Gmail credential named `Google Gmail OAuth2 API`.
4. Connect your Google account in the browser.
5. Copy the n8n API key and Gmail credential ID into `.env`.

Keep these defaults unless you know you need to change them:

```env
LOCAL_DOWNLOAD_PATH=./downloads/
N8N_CONTAINER_DOWNLOAD_PATH=/home/node/downloads/
```

## Run Everything

After `.env` contains a valid `N8N_API_KEY` and `GMAIL_CREDENTIAL_ID`, run:

```bash
docker compose up --build
```

The `my_project` container waits until n8n is healthy, creates the workflow, triggers it, and exits. n8n keeps running. Downloaded PDFs appear in:

```text
./downloads/
```

To rerun the downloader while n8n is already up:

```bash
docker compose run --rm my_project
```

## Switching From Manual Docker Run

If a manual container named `n8n` is already running, stop it once before using Compose:

```bash
docker stop n8n
```

Then start the Compose stack:

```bash
docker compose up --build
```

## Fresh Image Notes

The n8n Docker image itself is not modified by this project. A fresh n8n image works as long as it is started with:

```yaml
N8N_RESTRICT_FILE_ACCESS_TO: /home/node/downloads
```

and the local downloads folder is mounted to:

```text
/home/node/downloads
```

A fresh n8n data volume will not have your old API key or Gmail OAuth credential. Recreate those once in the n8n UI, update `.env`, and future Compose runs will reuse them from the `n8n_data` volume.
