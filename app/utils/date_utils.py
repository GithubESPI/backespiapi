import re

# Fonction pour formater une durée en minutes à partir d'une chaîne de caractères
def format_duration_to_minutes(duration_str):
    try:
        match = re.match(r'(?:(\d+)h)?(\d+)?', duration_str)
        if not match:
            raise ValueError("Invalid duration format")
        hours = int(match.group(1) or 0)
        minutes = int(match.group(2) or 0)
        return hours * 60 + minutes
    except (ValueError, TypeError) as e:
        raise ValueError(f"Error parsing duration: {e}")

# Fonction pour formater des minutes en une chaîne de caractères au format heures et minutes
def format_minutes_to_duration(minutes):
    if not isinstance(minutes, int) or minutes < 0:
        raise ValueError("Minutes should be a non-negative integer")
    hours, remaining_minutes = divmod(minutes, 60)
    return f"{hours}h{remaining_minutes:02d}" if hours else f"{remaining_minutes} minutes"

# Fonction pour sommer une liste de durées
def sum_durations(duration_list):
    if not all(isinstance(duration, int) and duration >= 0 for duration in duration_list):
        raise ValueError("All durations should be non-negative integers")
    total_minutes = sum(duration_list)
    return format_minutes_to_duration(total_minutes)
