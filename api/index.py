import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# Export app as WSGI application for Vercel
# Vercel automatically detects and uses this
__all__ = ['app']