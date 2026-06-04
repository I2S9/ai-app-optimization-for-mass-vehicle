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

    # Supabase HTTPS (port 443 — passe le pare-feu entreprise)
    supabase_url: str = ""
    supabase_service_key: str = ""

    # Supabase Postgres direct (port 6543 — souvent bloque)
    supabase_db_url: str = ""

    # supabase | default | databricks | postgres | auto
    data_backend: str = "supabase"
    local_data_dir: str = "../web/public/data"
    cors_origins: str = "http://127.0.0.1:5173,http://localhost:5173,http://127.0.0.1:8080"

    @property
    def use_supabase_rest(self) -> bool:
        if self.data_backend == "default":
            return False
        if self.data_backend == "databricks":
            return False
        return bool(self.supabase_url and self.supabase_service_key)

    @property
    def pg_url(self) -> str:
        return (self.supabase_db_url or "").strip()

    @property
    def use_postgres_direct(self) -> bool:
        if self.use_supabase_rest or not self.pg_url:
            return False
        if self.data_backend == "default":
            return False
        if self.data_backend == "databricks":
            return False
        return self.data_backend in ("supabase", "postgres", "local", "auto")

    @property
    def use_postgres_store(self) -> bool:
        return self.use_supabase_rest or self.use_postgres_direct

    @property
    def use_supabase(self) -> bool:
        return self.use_postgres_store

    @property
    def use_databricks(self) -> bool:
        if self.data_backend in ("default", "supabase", "postgres", "local"):
            return False
        if self.data_backend == "databricks":
            return bool(
                self.databricks_host
                and self.databricks_http_path
                and self.databricks_token
            )
        if self.use_postgres_store:
            return False
        return bool(
            self.databricks_host
            and self.databricks_http_path
            and self.databricks_token
        )

    @property
    def api_mode(self) -> str:
        if self.use_supabase_rest:
            return "supabase"
        if self.use_postgres_direct:
            if "supabase.co" in self.pg_url:
                return "supabase"
            return "postgres"
        if self.use_databricks:
            return "databricks"
        return "local"

    @property
    def use_remote_store(self) -> bool:
        return self.use_postgres_store or self.use_databricks

    @property
    def remote_only(self) -> bool:
        return self.data_backend == "supabase" and self.use_supabase_rest

    @property
    def local_data_path(self) -> Path:
        return (Path(__file__).resolve().parent.parent / self.local_data_dir).resolve()

    @property
    def cors_origin_list(self) -> list[str]:
        return [o.strip() for o in self.cors_origins.split(",") if o.strip()]


settings = Settings()
