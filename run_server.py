"""
Lance le serveur FastAPI depuis la racine du projet.
Usage: python run_server.py
"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run(
        "backend.main:app",
        host="127.0.0.1",
        port=8001,
        reload=True,
    )
