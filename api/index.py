import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# For Vercel serverless - export the Flask app as WSGI application
__all__ = ['app']