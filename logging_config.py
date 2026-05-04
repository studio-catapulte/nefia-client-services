"""Structured JSON logging with request_id propagation."""

import logging
import sys
from contextvars import ContextVar

import structlog

request_id_ctx: ContextVar[str] = ContextVar("request_id", default="-")


def _add_request_id(_, __, event_dict):
    event_dict["request_id"] = request_id_ctx.get()
    return event_dict


def configure_logging() -> None:
    logging.basicConfig(
        format="%(message)s",
        stream=sys.stdout,
        level=logging.INFO,
    )

    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            _add_request_id,
            structlog.processors.add_log_level,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.processors.format_exc_info,
            structlog.processors.JSONRenderer(),
        ],
        wrapper_class=structlog.make_filtering_bound_logger(logging.INFO),
        logger_factory=structlog.PrintLoggerFactory(),
        cache_logger_on_first_use=True,
    )


def get_logger(name: str = "nefia.client_services"):
    return structlog.get_logger(name)
