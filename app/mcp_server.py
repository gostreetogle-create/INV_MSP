"""MCP server exposing the Inventor COM bridge as tools for Claude.

Run standalone for testing:
    python mcp_server.py

Register with Claude Code:
    claude mcp add inventor -- python "<this file's full path>"
"""
from mcp.server.fastmcp import FastMCP

import core

mcp = FastMCP("inventor")


@mcp.tool()
def inventor_status() -> dict:
    """Check connection to Inventor and report what's open."""
    return core.get_status()


@mcp.tool()
def list_documents() -> list:
    """List documents currently open in Inventor."""
    return core.list_documents()


@mcp.tool()
def open_document(path: str) -> str:
    """Open an Inventor document (.ipt/.iam/.idw) by full file path."""
    return core.open_document(path)


@mcp.tool()
def list_parameters(document_name: str = None) -> list:
    """List model parameters for a document (active document if name omitted)."""
    return core.list_parameters(document_name)


@mcp.tool()
def get_parameter(name: str, document_name: str = None) -> dict:
    """Get a single named parameter's value/expression/units."""
    return core.get_parameter(name, document_name)


@mcp.tool()
def set_parameter(name: str, expression: str, document_name: str = None) -> dict:
    """Set a parameter's expression (e.g. '25 mm') and update the model."""
    return core.set_parameter(name, expression, document_name)


@mcp.tool()
def list_ilogic_rules(document_name: str = None) -> list:
    """List iLogic rule names present in a document."""
    return core.list_ilogic_rules(document_name)


@mcp.tool()
def run_ilogic_rule(rule_name: str, document_name: str = None) -> dict:
    """Run an existing iLogic rule by name in a document."""
    return core.run_ilogic_rule(rule_name, document_name)


@mcp.tool()
def save_document(document_name: str = None) -> dict:
    """Save a document."""
    return core.save_document(document_name)


@mcp.tool()
def export_document(target_path: str, document_name: str = None) -> dict:
    """Save/export a document to target_path (format inferred from extension, e.g. .pdf, .step)."""
    return core.export_document(target_path, document_name)


@mcp.tool()
def suggest_designation(level: str, product: int = None, module: int = None) -> str:
    """Suggest the next free PDM designation. level: 'product' | 'module' | 'part'."""
    return core.suggest_designation(level, product, module)


@mcp.tool()
def designation_exists(designation: str) -> bool:
    """Check whether a file with this PDM designation already exists in the project folder."""
    return core.designation_exists(designation)


@mcp.tool()
def create_assembly(designation: str, description: str = "") -> dict:
    """Create an empty assembly container ('Изделие' or 'Модуль') with PDM designation set."""
    return core.create_assembly(designation, description)


@mcp.tool()
def list_known_products() -> list:
    """Distinct product numbers already used by files in the project folder."""
    return core.list_known_products()


@mcp.tool()
def list_known_modules(product: int) -> list:
    """Distinct module numbers already used under a given product."""
    return core.list_known_modules(product)


@mcp.tool()
def list_tube_presets() -> list:
    """Standard tube cross-sections from catalog.xlsx."""
    return core.list_tube_presets()


@mcp.tool()
def list_sheet_thicknesses() -> list:
    """Standard sheet thicknesses from catalog.xlsx."""
    return core.list_sheet_thicknesses()


@mcp.tool()
def list_part_name_templates() -> list:
    """Common part-name patterns from catalog.xlsx for Description autocomplete."""
    return core.list_part_name_templates()


@mcp.tool()
def open_catalog() -> dict:
    """Open catalog.xlsx in Excel for editing."""
    return core.open_catalog()


@mcp.tool()
def list_materials() -> list:
    """List material names available in the default part template's library."""
    return core.list_materials()


@mcp.tool()
def create_sheet_part(designation: str, length_mm: float, width_mm: float, thickness_mm: float,
                       description: str = "", material_name: str = None, save_dir: str = None) -> dict:
    """Create a flat rectangular plate part ('Лист') with PDM designation and iProperties set."""
    return core.create_sheet_part(designation, length_mm, width_mm, thickness_mm,
                                   description, material_name, save_dir)


@mcp.tool()
def create_tube_part(designation: str, outer_width_mm: float, outer_height_mm: float,
                      wall_mm: float, length_mm: float, description: str = "",
                      material_name: str = None, save_dir: str = None) -> dict:
    """Create a hollow rectangular/square tube part ('Труба') with PDM designation and iProperties set."""
    return core.create_tube_part(designation, outer_width_mm, outer_height_mm, wall_mm, length_mm,
                                  description, material_name, save_dir)


if __name__ == "__main__":
    mcp.run()
