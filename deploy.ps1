# Script de deploy a Google Cloud Run
# Ejecutar desde el terminal de VS Code: .\deploy.ps1

$PROJECT_ID = "project-e7acaf8b-8b9b-4f52-ba4"
$SERVICE_NAME = "opencv-flask"
$REGION = "us-central1"

Write-Host "Desplegando $SERVICE_NAME a Cloud Run..." -ForegroundColor Cyan

gcloud run deploy $SERVICE_NAME `
    --source . `
    --region $REGION `
    --project $PROJECT_ID `
    --allow-unauthenticated `
    --platform managed

Write-Host "Deploy completado!" -ForegroundColor Green
Write-Host "URL: https://opencv-flask-xn43a3ze3a-uc.a.run.app" -ForegroundColor Yellow
