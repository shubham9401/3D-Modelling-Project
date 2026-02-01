<p align="center">
  <img src="https://img.shields.io/badge/SolidWorks-FF0000?style=for-the-badge&logo=dassaultsystemes&logoColor=white" alt="SolidWorks"/>
  <img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python"/>
  <img src="https://img.shields.io/badge/Powered%20by-Llama%203-blueviolet?style=for-the-badge" alt="Llama 3"/>
  <img src="https://img.shields.io/badge/Groq-LPU-orange?style=for-the-badge" alt="Groq"/>
</p>

<h1 align="center">🤖 SolidWorks AI Agent</h1>

<p align="center">
  <strong>Transform natural language into precision-engineered 3D CAD models</strong>
</p>

<p align="center">
  <a href="#-features">Features</a> •
  <a href="#-quick-start">Quick Start</a> •
  <a href="#-usage">Usage</a> •
  <a href="#-examples">Examples</a> •
  <a href="#-architecture">Architecture</a>
</p>

---

## 📋 Overview

The **SolidWorks AI Agent** is a powerful automation framework that bridges natural language and CAD modeling. Simply describe the 3D model you want to create in plain English, and the agent leverages **Llama 3** (via Groq's ultra-fast LPU inference) to generate and execute precise SolidWorks API commands.

> 💡 *"Create a 100x50mm rectangular plate with a 20mm hole in the center"* → **Fully modeled part in seconds**

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🗣️ **Natural Language Interface** | Describe your model in plain English—no scripting required |
| ⚡ **Lightning-Fast Inference** | Powered by Groq LPU for near-instant code generation |
| 🔧 **Comprehensive Toolset** | 30+ CAD operations including sketches, extrusions, cuts, revolves, threads, and more |
| 🛡️ **Safe Execution** | All AI-generated commands are validated before execution |
| 📐 **Precision Engineering** | Supports exact dimensions, coordinates, and geometric constraints |
| 🔩 **Hardware Modeling** | Create bolts, nuts, threaded holes, and mechanical components |
| 📊 **Smart Coordinate System** | Auto-calculates face positions and feature coordinates |

---

## 🚀 Quick Start

### Prerequisites

- **SolidWorks 2020+** — Installed and running
- **Python 3.10+** — With pip package manager
- **Groq API Key** — Free at [console.groq.com](https://console.groq.com)

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/3D-Modelling-Project.git
cd 3D-Modelling-Project

# Install dependencies
pip install groq pywin32 python-dotenv
```

### Configuration

Create a `.env` file in the project root:

```env
GROQ_API_KEY=gsk_your_api_key_here
GROQ_MODEL=llama-3.3-70b-versatile
```

Or set environment variables directly:

```powershell
# Windows PowerShell
$env:GROQ_API_KEY = "gsk_your_api_key_here"
```

---

## 🎮 Usage

### Interactive Mode

```bash
python main.py
```

1. **Describe your model** — Enter a natural language description
2. **Review the plan** — The AI generates a mission file with CAD operations
3. **Execute** — Confirm to build the model in SolidWorks

### Example Session

```
🔧 SOLIDWORKS AI AGENT
   Powered by Llama 3 (via Groq)
============================================================

Describe the 3D model you want to create:
(Example: 'Create a 100x50mm rectangular plate, 10mm thick')

Your request: Create a hex nut M10 size

-----------------------------------------
STEP 1: Generating CAD commands with AI...
-----------------------------------------
✅ Mission generated successfully!

✅ SolidWorks is connected.

Execute in SolidWorks now? (y/n): y
```

---

## 📝 Examples

| Description | Prompt |
|-------------|--------|
| Simple Box | `"Create a 100x50mm rectangular plate, 10mm thick"` |
| Plate with Hole | `"Create a 50x50mm plate with a 10mm diameter hole in the center"` |
| Sphere | `"Create a sphere with 30mm radius"` |
| Cone | `"Create a cone with 25mm base radius and 60mm height"` |
| Hex Nut | `"Create an M10 hex nut"` |
| Threaded Bolt | `"Create a hex head bolt M8x1.25, 40mm shaft length"` |
| Cup | `"Create a cup with 40mm radius, 100mm tall, 3mm walls"` |
| Filleted Box | `"Create a 50x50x30mm box with 5mm filleted edges"` |

---

## 📂 Architecture

```
3D-Modelling-Project/
│
├── main.py                 # Entry point & user interaction
├── llm_client.py           # Groq/Llama 3 API integration
├── system_prompt.py        # AI instructions & tool definitions
│
├── tools/                  # SolidWorks API wrappers
│   ├── solidworks_app.py   # Application connection & lifecycle
│   ├── part.py             # Part document management
│   ├── sketch.py           # 2D geometry creation
│   ├── feature.py          # 3D operations & features
│   └── assembly.py         # Assembly & mating operations
│
├── mcp_server/             # Mission execution engine
│   └── dispatcher.py       # Command dispatcher
│
└── mission.json            # Generated mission file
```

---

## ⚙️ Configuration

| Environment Variable | Default | Description |
|---------------------|---------|-------------|
| `GROQ_API_KEY` | *Required* | Your Groq API key |
| `GROQ_MODEL` | `llama-3.3-70b-versatile` | LLM model to use |

---

## ⚠️ Notes

- **Units**: All dimensions are in **millimeters (mm)** by default
- **SolidWorks**: Must be running before executing missions
- **Validation**: Review generated actions before executing complex models

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  <strong>Built with ❤️ for CAD automation</strong>
</p>

<p align="center">
  <sub>Powered by Groq's Lightning-Fast LPU Inference</sub>
</p>
