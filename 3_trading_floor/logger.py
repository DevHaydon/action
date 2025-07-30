from database import write_log


def log_error(name: str, message: str) -> None:
    """Log an error message for the given agent."""
    write_log(name.lower(), "error", message)


def log_exception(name: str, exc: Exception, context: str | None = None) -> None:
    """Log an exception with optional context."""
    msg = f"{context}: {exc}" if context else str(exc)
    log_error(name, msg)


def log_risk(name: str, message: str) -> None:
    """Log a risk-related incident for the given agent."""
    write_log(name.lower(), "risk", message)


def log_audit(name: str, message: str) -> None:
    """Log an audit trail message for the given agent."""
    write_log(name.lower(), "audit", message)
