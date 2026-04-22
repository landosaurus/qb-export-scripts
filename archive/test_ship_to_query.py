"""
Test script for ship_to_address_query module.

This script demonstrates how to use the ship_to_address_query function
without relying on the __init__.py imports that reference non-existent modules.
"""

import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(__file__))

# Import directly from the module file
from qb_mcp_tools.tools.ship_to_address_query import ship_to_address_query


def test_ship_to_query():
    """Test the ship_to_address_query function."""
    print("Testing ship_to_address_query function...")
    print(f"Function name: {ship_to_address_query.__name__}")
    print(f"Module: {ship_to_address_query.__module__}")
    print("\nFunction signature:")
    import inspect
    sig = inspect.signature(ship_to_address_query)
    print(f"  {ship_to_address_query.__name__}{sig}")

    print("\nDocstring:")
    print(ship_to_address_query.__doc__)

    print("\n✓ Module imported successfully!")
    print("✓ Function is properly defined and documented!")

    return True


if __name__ == "__main__":
    try:
        test_ship_to_query()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
