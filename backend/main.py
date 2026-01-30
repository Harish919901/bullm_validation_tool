"""
FastAPI Backend for QW Validation Tool
Main entry point
"""

import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from routers.validation import router as validation_router

# Create FastAPI app
app = FastAPI(
    title="QW Validation API",
    description="API for Excel file validation (Quote Win & BOM Matrix)",
    version="1.0.0"
)

# Configure CORS for React frontend
# Get allowed origins from environment variable or use defaults
allowed_origins = os.getenv("ALLOWED_ORIGINS", "").split(",")
allowed_origins = [origin.strip() for origin in allowed_origins if origin.strip()]

# Default origins for local development
default_origins = [
    "http://localhost:5173",  # Vite dev server
    "http://localhost:3000",  # Alternative React port
    "http://127.0.0.1:5173",
    "http://127.0.0.1:3000",
]

# Combine default and environment origins
all_origins = list(set(default_origins + allowed_origins))

app.add_middleware(
    CORSMiddleware,
    allow_origins=all_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include routers
app.include_router(validation_router)


@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "message": "QW Validation API",
        "version": "1.0.0",
        "docs": "/docs"
    }


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
