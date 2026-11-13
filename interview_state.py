"""Shared boundary between a temporary setup and a started interview."""

def interview_started(state):
    return all(isinstance(state.get(key), str) and state[key].strip()
               for key in ("company", "system", "system_type"))


def interview_title(state):
    if not interview_started(state):
        return None
    return " · ".join(" ".join(state[key].split()) for key in ("company", "system", "system_type"))
