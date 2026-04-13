"""
Training Data Collector — Auto-saves successful runs as future training examples.

After a successful SolidWorks execution + validation (score >= 80), the prompt
and resulting mission JSON are saved to training_examples/ for use in the next
fine-tuning round.
"""

import json
import os
import re
import shutil
from pathlib import Path
from datetime import datetime

TRAINING_DIR = Path(__file__).parent / "training_examples"


def save_successful_run(user_prompt, mission_file="mission.json"):
    """
    Save a validated, successful run as a training example.

    Args:
        user_prompt: The original user request text
        mission_file: Path to the mission JSON that was executed successfully
    """
    if not user_prompt or not user_prompt.strip():
        return

    # Ensure directory exists
    TRAINING_DIR.mkdir(exist_ok=True)

    # Create a safe filename from the prompt
    safe_name = _prompt_to_filename(user_prompt)

    # Check for duplicates
    txt_path = TRAINING_DIR / f"{safe_name}.txt"
    json_path = TRAINING_DIR / f"{safe_name}.json"

    if txt_path.exists():
        # Append timestamp to avoid overwriting
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = f"{safe_name}_{ts}"
        txt_path = TRAINING_DIR / f"{safe_name}.txt"
        json_path = TRAINING_DIR / f"{safe_name}.json"

    # Load the mission JSON
    mission_path = Path(mission_file)
    if not mission_path.is_absolute():
        mission_path = Path(__file__).parent / mission_file

    if not mission_path.exists():
        print(f"⚠️  Could not save training example: {mission_file} not found")
        return

    try:
        with open(mission_path, "r", encoding="utf-8") as f:
            actions = json.load(f)

        # Validate it's a list of actions
        if not isinstance(actions, list) or len(actions) == 0:
            return

        # Save the pair
        txt_path.write_text(user_prompt.strip(), encoding="utf-8")
        json_path.write_text(json.dumps(actions, indent=4), encoding="utf-8")

        print(f"📦 Training example saved: {safe_name}")
        print(f"   Prompt:  {txt_path}")
        print(f"   Actions: {json_path}")

    except (json.JSONDecodeError, IOError) as e:
        print(f"⚠️  Could not save training example: {e}")


def _prompt_to_filename(prompt):
    """Convert a prompt to a safe filename."""
    # Lowercase, replace spaces with underscores
    name = prompt.lower().strip()
    # Remove common prefixes
    for prefix in ["create a ", "create an ", "make a ", "make an ", "design a ", "design an "]:
        if name.startswith(prefix):
            name = name[len(prefix):]
            break
    # Keep only alphanumeric and spaces
    name = re.sub(r'[^a-z0-9\s]', '', name)
    # Replace spaces with underscores, collapse multiples
    name = re.sub(r'\s+', '_', name.strip())
    # Truncate
    return name[:60]


def get_example_count():
    """Return the number of training examples currently saved."""
    if not TRAINING_DIR.exists():
        return 0
    return len(list(TRAINING_DIR.glob("*.txt")))
