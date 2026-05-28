"""WGHT Vehicle Mass — API settings."""

from pathlib import Path
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    databricks_host: str = ""
    databricks_http_path: str = ""
    databricks_token: str = ""
    databricks_catalog: str = "main"
    databricks_schema: str = "vehicle_mass"

    data_backend: str = "auto"  # default | databricks | auto
    local_data_dir: str = "../web/public/data"
    cors_origins: str = "http://127.0.0.1:5173,http://localhost:5173"

    @property
    def use_databricks(self) -> bool:
        if self.data_backend == "default":
            return False
        if self.data_backend == "databricks":
            return True
        return bool(
            self.databricks_host
            and self.databricks_http_path
            and self.databricks_token
        )

    @property
    def local_data_path(self) -> Path:
        return (Path(__file__).resolve().parent.parent / self.local_data_dir).resolve()

    @property
    def cors_origin_list(self) -> list[str]:
        return [o.strip() for o in self.cors_origins.split(",") if o.strip()]


settings = Settings()
