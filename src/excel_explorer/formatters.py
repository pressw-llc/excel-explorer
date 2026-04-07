def format_metadata(meta: dict) -> str:
    """Format metadata dict as YAML-ish header lines."""
    lines = []
    for key, value in meta.items():
        if isinstance(value, bool):
            lines.append(f"{key}: {'true' if value else 'false'}")
        else:
            lines.append(f"{key}: {value}")
    return "\n".join(lines)


def format_output(meta: dict, body: str) -> str:
    """Format complete command output with metadata header and body."""
    header = format_metadata(meta)
    return f"{header}\n---\n{body}"
