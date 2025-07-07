from . import server
import asyncio
import argparse

def main():
    """Main entry point for the package."""
    parser = argparse.ArgumentParser(description='Powerpoint MCP Server')
    parser.add_argument('--folder-path',
                       default="/users/russellashby/decks/",
                       help="Folder to store completed decks in.")
    parser.add_argument('--owui-url',

                       help="URL of the Open-WebUI server to upload completed decks to.")
    parser.add_argument('--owui-token',
                       help="Token for the Open-WebUI server to upload completed decks to.")
    args = parser.parse_args()
    asyncio.run(server.main(args.folder_path, args.owui_url, args.owui_token))

# Optionally expose other important items at package level
__all__ = ['main', 'server']