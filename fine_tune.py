"""
Fine-Tuning Script for Databricks GPT-OSS-120B

Validates, uploads training data, and launches a fine-tuning job
on Azure Databricks Mosaic AI Model Training.

Usage:
    python fine_tune.py                    # Full pipeline: validate → upload → train
    python fine_tune.py --validate-only    # Just validate the JSONL
    python fine_tune.py --status           # Check status of an existing job

Prerequisites:
    pip install databricks-sdk
"""

import json
import os
import sys
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

TRAINING_FILE = "training_data.jsonl"
PROJECT_DIR = Path(__file__).parent

# Databricks config from .env
DATABRICKS_HOST = os.environ.get("LLM_BASE_URL", "").replace("/serving-endpoints", "")
DATABRICKS_TOKEN = os.environ.get("LLM_API_KEY", "")
BASE_MODEL = os.environ.get("LLM_MODEL", "databricks-gpt-oss-120b")

# Fine-tuning hyperparameters
FINETUNE_CONFIG = {
    "n_epochs": 3,
    "batch_size": 4,
    "learning_rate_multiplier": 2.0,
    "warmup_ratio": 0.1,
}


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_training_data(filepath):
    """Validate the JSONL training file before uploading."""
    filepath = Path(filepath)
    if not filepath.exists():
        print(f"❌ Training file not found: {filepath}")
        print("   Run 'python generate_training_data.py' first.")
        return False

    total = 0
    errors = []

    with open(filepath, "r", encoding="utf-8") as f:
        for line_num, line in enumerate(f, 1):
            line = line.strip()
            if not line:
                continue
            total += 1

            try:
                row = json.loads(line)
            except json.JSONDecodeError as e:
                errors.append(f"Line {line_num}: Invalid JSON - {e}")
                continue

            if "messages" not in row:
                errors.append(f"Line {line_num}: Missing 'messages' key")
                continue

            roles = [m.get("role") for m in row["messages"]]
            if roles != ["system", "user", "assistant"]:
                errors.append(f"Line {line_num}: Bad role sequence: {roles}")

            # Verify assistant content is valid JSON
            try:
                assistant_content = row["messages"][2]["content"]
                actions = json.loads(assistant_content)
                if not isinstance(actions, list):
                    errors.append(f"Line {line_num}: Assistant content is not a list")
            except (json.JSONDecodeError, IndexError):
                errors.append(f"Line {line_num}: Assistant content is not valid JSON")

    if errors:
        print(f"❌ Validation failed with {len(errors)} errors:")
        for e in errors[:10]:
            print(f"  {e}")
        return False

    print(f"✅ Validation passed: {total} examples")
    if total < 50:
        print(f"⚠️  Warning: {total} examples is below the recommended minimum of 50")
    return True


# ---------------------------------------------------------------------------
# Databricks Fine-Tuning via Mosaic AI
# ---------------------------------------------------------------------------

def finetune_databricks(filepath):
    """Launch fine-tuning on Databricks Mosaic AI Model Training."""
    try:
        from databricks.sdk import WorkspaceClient
        from databricks.sdk.service.serving import EndpointCoreConfigInput
    except ImportError:
        print("❌ databricks-sdk not installed.")
        print("   Install it with: pip install databricks-sdk")
        return False

    filepath = Path(filepath)
    if not validate_training_data(filepath):
        return False

    # Initialize Databricks client
    host = DATABRICKS_HOST
    token = DATABRICKS_TOKEN

    if not host or not token:
        print("❌ Databricks credentials not configured.")
        print("   Set LLM_BASE_URL and LLM_API_KEY in your .env file.")
        return False

    print(f"\n{'='*50}")
    print(f"DATABRICKS FINE-TUNING")
    print(f"{'='*50}")
    print(f"Host:       {host}")
    print(f"Base Model: {BASE_MODEL}")
    print(f"Data File:  {filepath}")
    print(f"Epochs:     {FINETUNE_CONFIG['n_epochs']}")
    print(f"Batch Size: {FINETUNE_CONFIG['batch_size']}")

    try:
        w = WorkspaceClient(host=host, token=token)

        # Step 1: Upload training data to DBFS
        print("\n📤 Uploading training data to DBFS...")
        dbfs_path = f"/FileStore/fine-tuning/{filepath.name}"

        with open(filepath, "rb") as f:
            w.dbfs.put(dbfs_path, f, overwrite=True)
        print(f"   Uploaded to: dbfs:{dbfs_path}")

        # Step 2: Create fine-tuning run
        print("\n🚀 Launching fine-tuning job...")
        run_name = f"solidworks-agent-ft-{int(time.time())}"

        # Use Mosaic AI Model Training API
        from databricks.sdk.service.catalog import VolumeType
        run = w.model_training.create(
            model=BASE_MODEL,
            train_data_path=f"dbfs:{dbfs_path}",
            register_to=f"solidworks_agent_finetuned",
            training_duration=f"{FINETUNE_CONFIG['n_epochs']} epochs",
            learning_rate=FINETUNE_CONFIG['learning_rate_multiplier'],
            context_length=4096,
        )

        print(f"   Run created: {run_name}")
        print(f"   Run ID: {run.run_id if hasattr(run, 'run_id') else 'N/A'}")

        # Step 3: Monitor progress
        print("\n⏳ Monitoring progress...")
        _monitor_databricks_run(w, run)

        return True

    except Exception as e:
        print(f"\n❌ Databricks fine-tuning failed: {e}")
        print(f"\n💡 Alternative: Use the OpenAI-compatible fine-tuning approach.")
        print(f"   See the manual instructions below.\n")
        _print_manual_instructions(filepath)
        return False


def _monitor_databricks_run(client, run):
    """Monitor a Databricks fine-tuning run until completion."""
    try:
        run_id = run.run_id if hasattr(run, 'run_id') else None
        if not run_id:
            print("   Could not get run ID. Check Databricks UI for progress.")
            return

        while True:
            status = client.model_training.get(run_id)
            state = status.state if hasattr(status, 'state') else "UNKNOWN"
            print(f"   Status: {state}")

            if state in ("COMPLETED", "FAILED", "CANCELLED"):
                break

            time.sleep(30)

        if state == "COMPLETED":
            model_name = status.registered_model if hasattr(status, 'registered_model') else "solidworks_agent_finetuned"
            print(f"\n🎉 Fine-tuning complete!")
            print(f"   Model: {model_name}")
            print(f"\n   Add to .env:")
            print(f"   LLM_FINETUNED_MODEL={model_name}")
            _update_env_file(model_name)
        else:
            print(f"\n❌ Fine-tuning {state.lower()}")

    except Exception as e:
        print(f"   Monitoring error: {e}")
        print("   Check Databricks UI for job status.")


# ---------------------------------------------------------------------------
# OpenAI-Compatible Fine-Tuning (Alternative)
# ---------------------------------------------------------------------------

def finetune_openai_compatible(filepath):
    """
    Fine-tune using the OpenAI API format.
    Works with OpenAI, or any OpenAI-compatible fine-tuning endpoint.
    """
    try:
        from openai import OpenAI
    except ImportError:
        print("❌ openai package not installed. Run: pip install openai")
        return False

    filepath = Path(filepath)
    if not validate_training_data(filepath):
        return False

    api_key = os.environ.get("LLM_API_KEY", "")
    base_url = os.environ.get("LLM_BASE_URL", "")
    model = os.environ.get("LLM_MODEL", "")

    if not api_key:
        print("❌ LLM_API_KEY not set in .env")
        return False

    client = OpenAI(api_key=api_key, base_url=base_url)

    print(f"\n{'='*50}")
    print(f"OPENAI-COMPATIBLE FINE-TUNING")
    print(f"{'='*50}")
    print(f"Base URL:   {base_url}")
    print(f"Base Model: {model}")

    try:
        # Step 1: Upload file
        print("\n📤 Uploading training file...")
        with open(filepath, "rb") as f:
            file_obj = client.files.create(file=f, purpose="fine-tune")
        print(f"   File ID: {file_obj.id}")

        # Step 2: Create fine-tuning job
        print("\n🚀 Creating fine-tuning job...")
        job = client.fine_tuning.jobs.create(
            training_file=file_obj.id,
            model=model,
            hyperparameters={
                "n_epochs": FINETUNE_CONFIG["n_epochs"],
                "batch_size": FINETUNE_CONFIG["batch_size"],
                "learning_rate_multiplier": FINETUNE_CONFIG["learning_rate_multiplier"],
            },
        )
        print(f"   Job ID: {job.id}")
        print(f"   Status: {job.status}")

        # Step 3: Monitor
        print("\n⏳ Monitoring (Ctrl+C to stop monitoring, job continues)...")
        while True:
            job = client.fine_tuning.jobs.retrieve(job.id)
            print(f"   Status: {job.status}")

            if job.status in ("succeeded", "failed", "cancelled"):
                break
            time.sleep(30)

        if job.status == "succeeded":
            model_name = job.fine_tuned_model
            print(f"\n🎉 Fine-tuning complete!")
            print(f"   Fine-tuned model: {model_name}")
            _update_env_file(model_name)
        else:
            print(f"\n❌ Fine-tuning {job.status}")
            if hasattr(job, 'error') and job.error:
                print(f"   Error: {job.error}")

        return job.status == "succeeded"

    except Exception as e:
        print(f"\n❌ Fine-tuning failed: {e}")
        import traceback
        traceback.print_exc()
        return False


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _update_env_file(model_name):
    """Add or update LLM_FINETUNED_MODEL in .env file."""
    env_path = PROJECT_DIR / ".env"

    if env_path.exists():
        content = env_path.read_text(encoding="utf-8")
    else:
        content = ""

    # Update or add the line
    key = "LLM_FINETUNED_MODEL"
    new_line = f"{key}={model_name}"

    if key in content:
        # Replace existing line
        import re
        content = re.sub(rf'^{key}=.*$', new_line, content, flags=re.MULTILINE)
    else:
        content = content.rstrip() + f"\n{new_line}\n"

    env_path.write_text(content, encoding="utf-8")
    print(f"\n📝 Updated .env with: {new_line}")


def _print_manual_instructions(filepath):
    """Print manual fine-tuning instructions."""
    print("=" * 50)
    print("MANUAL FINE-TUNING INSTRUCTIONS")
    print("=" * 50)
    print(f"""
If the automated approach fails, you can fine-tune manually:

1. DATABRICKS (your current setup):
   - Go to your Databricks workspace: {DATABRICKS_HOST}
   - Navigate to: Machine Learning → Experiments → Create
   - Upload: {filepath.absolute()}
   - Select base model: {BASE_MODEL}
   - Configure: epochs={FINETUNE_CONFIG['n_epochs']}, batch_size={FINETUNE_CONFIG['batch_size']}
   - Start training
   - After completion, update .env: LLM_FINETUNED_MODEL=<model_name>

2. OPENAI (alternative):
   - Get API key from: https://platform.openai.com
   - Update .env: LLM_BASE_URL=https://api.openai.com/v1
   - Run: python fine_tune.py --openai

3. LOCAL (Llama 3 + LoRA):
   - pip install unsloth
   - See: https://github.com/unslothai/unsloth
   - Training data is ready in: {filepath.absolute()}
""")


def check_status():
    """Check the status of a running fine-tuning job."""
    print("📊 Checking fine-tuning job status...")

    try:
        from openai import OpenAI
        client = OpenAI(
            api_key=os.environ.get("LLM_API_KEY", ""),
            base_url=os.environ.get("LLM_BASE_URL", ""),
        )

        jobs = client.fine_tuning.jobs.list(limit=5)
        if not jobs.data:
            print("   No fine-tuning jobs found.")
            return

        for job in jobs.data:
            print(f"\n   Job: {job.id}")
            print(f"   Model: {job.model}")
            print(f"   Status: {job.status}")
            if job.fine_tuned_model:
                print(f"   Fine-tuned: {job.fine_tuned_model}")
            if hasattr(job, 'created_at'):
                print(f"   Created: {job.created_at}")

    except Exception as e:
        print(f"   Error: {e}")
        print("   Check Databricks UI for job status.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Load .env
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except ImportError:
        pass

    filepath = PROJECT_DIR / TRAINING_FILE

    print("🔧 SolidWorks AI Agent — Fine-Tuning Pipeline")
    print("=" * 50)

    if "--validate-only" in sys.argv:
        validate_training_data(filepath)
        return

    if "--status" in sys.argv:
        check_status()
        return

    if "--openai" in sys.argv:
        finetune_openai_compatible(filepath)
        return

    # Default: try Databricks first
    if not filepath.exists():
        print(f"❌ Training data not found: {filepath}")
        print("   Run 'python generate_training_data.py' first.")
        return

    print(f"\nTraining data: {filepath}")

    # Count examples
    with open(filepath, "r") as f:
        count = sum(1 for line in f if line.strip())
    print(f"Examples: {count}")

    if count < 10:
        print("❌ Too few examples. Need at least 10, recommend 50+.")
        return

    print(f"\nSelect fine-tuning backend:")
    print(f"  [1] Databricks Mosaic AI (current provider)")
    print(f"  [2] OpenAI-compatible API")
    print(f"  [q] Quit")

    choice = input("\nChoice (1/2/q): ").strip()

    if choice == "1":
        finetune_databricks(filepath)
    elif choice == "2":
        finetune_openai_compatible(filepath)
    else:
        print("Cancelled.")


if __name__ == "__main__":
    main()
