from shapes.Cube_template import CubeBuilder
from shapes.Cuboid_template import CuboidBuilder
from shapes.Sphere_template import SphereBuilder
from shapes.Cone_template import ConeBuilder
from shapes.Cylinder_template import CylinderBuilder
from shapes.Wedge_template import WedgeBuilder

# 1. The Registry
BUILDERS = {
    "cube": CubeBuilder(),
    "cuboid": CuboidBuilder(),
    "sphere": SphereBuilder(),
    "cone": ConeBuilder(),
    "cylinder": CylinderBuilder(),
    "wedge": WedgeBuilder(),
}

# 2. The Execution Logic
def execute(part, data):
    shape_type = data["shape"]
    
    if shape_type in BUILDERS:
        BUILDERS[shape_type].build(part, data)
    else:
        print(f"Error: No builder found for shape type: {shape_type}")