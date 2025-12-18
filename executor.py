from shapes.Cube_template import CubeBuilder
from shapes.Cuboid_template import CuboidBuilder
from shapes.Sphere_template import SphereBuilder

# 1. The Registry
BUILDERS = {
    "cube": CubeBuilder(),
    "cuboid": CuboidBuilder(),
    "sphere": SphereBuilder(),
}

# 2. The Execution Logic
def execute(part, data):
    shape_type = data["shape"]
    
    if shape_type in BUILDERS:
        # Passes the part file and the specific data row to the builder
        BUILDERS[shape_type].build(part, data)
    else:
        print(f"Error: No builder found for shape type: {shape_type}")