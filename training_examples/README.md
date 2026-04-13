# Training Examples Directory

This directory stores verified `(prompt, JSON)` pairs for fine-tuning.

## How Examples Are Added

1. **Automatically** — When you run a CREATE flow and the validator scores it ≥ 80, the prompt + mission JSON are saved here.
2. **Manually** — You can add your own examples by creating paired files:
   - `my_shape.txt` — contains the natural language prompt
   - `my_shape.json` — contains the correct JSON array of tool calls

## File Format

**Prompt file** (`.txt`):
```
Create a 50mm cube with a 10mm hole through the top
```

**Action file** (`.json`):
```json
[
    {"tool": "create_part", "args": {}},
    {"tool": "create_sketch", "args": {"plane": "Top"}},
    ...
]
```

## Rebuilding Training Data

After adding new examples here, regenerate the training JSONL:
```
python generate_training_data.py
```

Then re-run fine-tuning:
```
python fine_tune.py
```
