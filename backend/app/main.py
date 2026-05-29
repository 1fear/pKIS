from fastapi import APIRouter, Depends, FastAPI, Header, HTTPException, status

from .schemas import HealthResponse, ImportCreate, ScanCreate
from .settings import APP_VERSION, load_settings


settings = load_settings()

app = FastAPI(
    title="TakSklad Backend API",
    version=APP_VERSION,
    docs_url="/docs",
    redoc_url="/redoc",
)


def require_service_token(authorization: str | None = Header(default=None)):
    if not settings.api_auth_enabled:
        return
    expected = f"Bearer {settings.api_token}"
    if authorization != expected:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid service token",
        )


@app.get("/health", response_model=HealthResponse)
def health():
    return {
        "status": "ok",
        "service": settings.service_name,
        "version": APP_VERSION,
        "environment": settings.environment,
    }


api = APIRouter(prefix="/api/v1", dependencies=[Depends(require_service_token)])


def not_implemented(feature):
    raise HTTPException(
        status_code=status.HTTP_501_NOT_IMPLEMENTED,
        detail=f"{feature} is defined in the MVP contract but not implemented yet.",
    )


@api.get("/orders/active")
def list_active_orders():
    not_implemented("Active orders")


@api.post("/scans")
def create_scan(payload: ScanCreate):
    not_implemented("Scan ingestion")


@api.post("/orders/{order_id}/complete")
def complete_order(order_id: str):
    not_implemented("Order completion")


@api.post("/imports")
def create_import(payload: ImportCreate):
    not_implemented("Excel import")


@api.get("/imports")
def list_imports():
    not_implemented("Import history")


@api.get("/reports/day")
def day_report(report_date: str | None = None):
    not_implemented("Day report")


app.include_router(api)
