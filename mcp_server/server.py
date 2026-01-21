import sys
import os
from mcp.server.fastmcp import FastMCP

# --- 1. PATH SETUP (CRITICAL) ---
# This tells Python to look in the main project folder for 'tools'
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.append(project_root)

# Now we can import your team's code
try:
    from tools import part, sketch, feature, assembly
    print("✅ SolidWorks Tools Imported Successfully")
except ImportError as e:
    print(f"❌ Error Importing Tools: {e}")
    print(f"   Looking in: {project_root}")
    sys.exit(1)

# --- 2. CREATE SERVER ---
mcp = FastMCP("SolidWorks Agent")

# --- 3. REGISTER TOOLS ---
# We wrap Shubham's functions so the AI can see them.

# -- PART TOOLS --
@mcp.tool()
def create_new_part():
    """Creates a new blank SolidWorks part."""
    return part.create_part()

# -- SKETCH TOOLS --
@mcp.tool()
def create_sketch(plane: str):
    """Starts a sketch. Plane must be 'Front', 'Top', or 'Right'."""
    return sketch.create_sketch(plane)

@mcp.tool()
def draw_rectangle(width: float, height: float):
    """Draws a center rectangle (units: mm)."""
    return sketch.draw_rectangle(width, height)

@mcp.tool()
def draw_circle(radius: float):
    """Draws a circle (units: mm)."""
    return sketch.draw_circle(radius)

@mcp.tool()
def draw_slot(length: float, width: float):
    """Draws a slot with center-to-center length (units: mm)."""
    return sketch.draw_slot(length, width)

@mcp.tool()
def draw_polygon(sides: int, radius: float):
    """Draws a regular polygon (e.g., Hexagon)."""
    return sketch.draw_polygon(sides, radius)

@mcp.tool()
def validate_sketch():
    """Checks if the sketch is closed and ready for extrusion."""
    return sketch.validate_closed_profile()

# -- FEATURE TOOLS --
@mcp.tool()
def extrude_boss(depth: float):
    """Extrudes the current sketch to create a solid (units: mm)."""
    return feature.extrude(depth)

@mcp.tool()
def cut_extrude(depth: float):
    """Cuts into the material (units: mm)."""
    return feature.cut_extrude(depth)

@mcp.tool()
def fillet_edges(radius: float):
    """Applies a fillet to selected edges."""
    return feature.fillet(radius)

# --- 4. START SERVER ---
if __name__ == "__main__":
    print("🚀 SolidWorks MCP Server Running...")
    mcp.run()