"""Cola AFIP: Streamlit encola; worker local ejecuta (dry-run o Playwright)."""

from .jobs import (
    ACTIONS,
    Job,
    create_job,
    enqueue_job,
    list_jobs,
    jobs_root,
)
from .registry import badge_label, ensure_cuit_registered, list_cuits

__all__ = [
    "ACTIONS",
    "Job",
    "create_job",
    "enqueue_job",
    "list_jobs",
    "jobs_root",
    "badge_label",
    "ensure_cuit_registered",
    "list_cuits",
]
