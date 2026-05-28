"""Verifies the FastAPI app can load and all routes are registered."""
import sys
sys.path.insert(0, r'c:\Users\naniv\OneDrive\Desktop\backend')

from main import app

print(f"App title: {app.title} v{app.version}")
print("\nRegistered routes:")
for route in app.routes:
    if hasattr(route, 'methods') and route.methods:
        print(f"  {list(route.methods)[0]:6} {route.path}")

print("\nApp loaded successfully.")
print("To start: uvicorn main:app --host 0.0.0.0 --port 8000 --reload")
print("Swagger:  http://localhost:8000/docs")
