"""
PlastiSense AI — FastAPI Backend
===============================
Serves the web frontend and exposes prediction APIs for both
the 1D-CNN spectral model and the ResNet-18 image model.

Run:
    cd app/
    python -m uvicorn server:app --host 0.0.0.0 --port 8000 --reload
"""

import os
import io
import time
import traceback

import numpy as np
import pandas as pd
import torch
import torch.nn as nn
from torchvision import models, transforms
from PIL import Image

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

# ── Paths (relative to this file) ────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)

CNN_MODEL_PATH    = os.path.join(PROJECT_DIR, "1d cnn", "1d_cnn_spectral_model.pth")
RESNET_MODEL_PATH = os.path.join(PROJECT_DIR, "restnet_model", "resnet18_microplastic.pth")
STATIC_DIR        = os.path.join(BASE_DIR, "static")

CLASSES = ['PE', 'PET', 'PP', 'PS']
DEVICE  = torch.device("cuda" if torch.cuda.is_available() else "cpu")

# ═════════════════════════════════════════════════════════════════
# Model Definitions
# ═════════════════════════════════════════════════════════════════

class Microplastic1DCNN(nn.Module):
    """1D-CNN for hyperspectral polymer classification.
    Input: (Batch, 1, 256) → Output: (Batch, 4)
    """
    def __init__(self, num_classes: int = 4):
        super().__init__()
        self.features = nn.Sequential(
            nn.Conv1d(1, 16, kernel_size=5, padding=2),
            nn.ReLU(),
            nn.MaxPool1d(2),
            nn.Conv1d(16, 32, kernel_size=5, padding=2),
            nn.ReLU(),
            nn.MaxPool1d(2),
        )
        self.classifier = nn.Sequential(
            nn.Flatten(),
            nn.Linear(32 * 64, 128),
            nn.ReLU(),
            nn.Dropout(0.5),
            nn.Linear(128, num_classes),
        )

    def forward(self, x):
        return self.classifier(self.features(x))


def build_resnet18(num_classes: int = 4):
    """Build ResNet-18 with the same head used during training."""
    model = models.resnet18(weights=None)
    num_ftrs = model.fc.in_features
    model.fc = nn.Sequential(
        nn.Dropout(0.5),
        nn.Linear(num_ftrs, num_classes),
    )
    return model


# ═════════════════════════════════════════════════════════════════
# Load Models
# ═════════════════════════════════════════════════════════════════

cnn_model = None
resnet_model = None

def load_models():
    """Load both models into memory."""
    global cnn_model, resnet_model

    # 1D-CNN
    try:
        cnn_model = Microplastic1DCNN(num_classes=len(CLASSES))
        state = torch.load(CNN_MODEL_PATH, map_location=DEVICE, weights_only=True)
        cnn_model.load_state_dict(state)
        cnn_model.to(DEVICE)
        cnn_model.eval()
        print(f"✓ 1D-CNN loaded from: {CNN_MODEL_PATH}")
    except Exception as e:
        print(f"✗ 1D-CNN failed to load: {e}")
        cnn_model = None

    # ResNet-18
    try:
        resnet_model = build_resnet18(num_classes=len(CLASSES))
        state = torch.load(RESNET_MODEL_PATH, map_location=DEVICE, weights_only=True)
        resnet_model.load_state_dict(state)
        resnet_model.to(DEVICE)
        resnet_model.eval()
        print(f"✓ ResNet-18 loaded from: {RESNET_MODEL_PATH}")
    except Exception as e:
        print(f"✗ ResNet-18 failed to load: {e}")
        resnet_model = None


# ═════════════════════════════════════════════════════════════════
# Image Preprocessing (same as training)
# ═════════════════════════════════════════════════════════════════

IMG_SIZE = 64

image_transform = transforms.Compose([
    transforms.Resize((IMG_SIZE, IMG_SIZE)),
    transforms.ToTensor(),
    transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225]),
])


# ═════════════════════════════════════════════════════════════════
# FastAPI App
# ═════════════════════════════════════════════════════════════════

app = FastAPI(title="PlastiSense AI", version="1.0.0")


@app.on_event("startup")
async def startup():
    load_models()


# ── Serve Frontend ────────────────────────────────────────────────
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
async def serve_frontend():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


# ── Health Check ──────────────────────────────────────────────────
@app.get("/api/health")
async def health_check():
    return {
        "status": "ok",
        "device": str(DEVICE),
        "resnet_loaded": resnet_model is not None,
        "cnn_loaded": cnn_model is not None,
    }


# ── Image Prediction (ResNet-18) ─────────────────────────────────
@app.post("/api/predict/image")
async def predict_image(file: UploadFile = File(...)):
    if resnet_model is None:
        raise HTTPException(status_code=503, detail="ResNet-18 model not loaded")

    try:
        contents = await file.read()
        img = Image.open(io.BytesIO(contents)).convert("RGB")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid image file")

    # Preprocess
    tensor = image_transform(img).unsqueeze(0).to(DEVICE)

    # Inference
    t0 = time.perf_counter()
    with torch.no_grad():
        outputs = resnet_model(tensor)
        probs = torch.nn.functional.softmax(outputs, dim=1)[0]
    elapsed_ms = (time.perf_counter() - t0) * 1000

    confidence, pred_idx = torch.max(probs, 0)
    pred_class = CLASSES[pred_idx.item()]

    return {
        "predicted_class": pred_class,
        "confidence": round(confidence.item(), 4),
        "time_ms": round(elapsed_ms, 2),
        "probabilities": {
            cls: round(probs[i].item(), 4) for i, cls in enumerate(CLASSES)
        },
        "model": "ResNet-18",
    }


# ── Spectral Prediction (1D-CNN) ─────────────────────────────────
@app.post("/api/predict/spectral")
async def predict_spectral(file: UploadFile = File(...)):
    if cnn_model is None:
        raise HTTPException(status_code=503, detail="1D-CNN model not loaded")

    try:
        contents = await file.read()
        df = pd.read_csv(io.BytesIO(contents))
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid CSV file")

    # Extract spectral features
    # The training CSV has labels in cols 0-2, features from col 3 onward (256 channels)
    # For inference we try to be flexible:
    #   - If the CSV has > 256 columns, take columns from index 3 onward (same as training)
    #   - If the CSV has exactly 256 columns, use all of them directly
    #   - If a single row is provided, use that; otherwise use the first row
    try:
        if df.shape[1] > 256:
            features = df.iloc[0, 3:3+256].values.astype(np.float32)
        elif df.shape[1] == 256:
            features = df.iloc[0].values.astype(np.float32)
        else:
            # Try to find numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) >= 256:
                features = df[numeric_cols[:256]].iloc[0].values.astype(np.float32)
            else:
                raise ValueError(
                    f"Expected at least 256 numeric columns, got {len(numeric_cols)}. "
                    f"CSV should contain 256 reflectance channels."
                )
    except (IndexError, ValueError) as e:
        raise HTTPException(
            status_code=400,
            detail=f"Could not extract spectral data: {str(e)}"
        )

    # Build tensor: (1, 1, 256)
    x_tensor = torch.tensor(features).unsqueeze(0).unsqueeze(0).to(DEVICE)

    # Inference
    t0 = time.perf_counter()
    with torch.no_grad():
        outputs = cnn_model(x_tensor)
        probs = torch.nn.functional.softmax(outputs, dim=1)[0]
    elapsed_ms = (time.perf_counter() - t0) * 1000

    confidence, pred_idx = torch.max(probs, 0)
    pred_class = CLASSES[pred_idx.item()]

    return {
        "predicted_class": pred_class,
        "confidence": round(confidence.item(), 4),
        "time_ms": round(elapsed_ms, 2),
        "probabilities": {
            cls: round(probs[i].item(), 4) for i, cls in enumerate(CLASSES)
        },
        "model": "1D-CNN",
        "csv_shape": list(df.shape),
    }
