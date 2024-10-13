import logging
from app.api.endpoints import uploads, importBulletin
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

# Configurer le logger pour la production
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# Ajouter la middleware CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://bulletin.groupe-espi.fr"],  # Remplacer par l'URL de ton frontend
    allow_credentials=True,
    allow_methods=["*"],  # Autorise toutes les méthodes HTTP (GET, POST, PUT, DELETE, etc.)
    allow_headers=["*"],  # Autorise tous les en-têtes
)

# Inclusion des routes des différents modules
app.include_router(uploads.router, prefix="", tags=["uploads"])  # Uploads sans préfixe
app.include_router(importBulletin.router, prefix="/importBulletins", tags=["importBulletins"])


@app.get("/")
def read_root():
    return {"message": "Bonjour"}

# Pour lancer l'application en production, utilisez la commande suivante :
# gunicorn -w 4 -k uvicorn.workers.UvicornWorker app.main:app