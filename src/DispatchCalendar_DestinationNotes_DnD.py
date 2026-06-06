"""Pass a dropped Excel file name to the DestinationNotes command script."""

from pathlib import Path
import subprocess
import sys


COMMAND_SCRIPT_NAME = "DispatchCalendar_DestinationNotes_Cmd.py"


def get_command_script_path() -> Path:
    """Return the command script path located next to this script."""
    return Path(__file__).resolve().with_name(COMMAND_SCRIPT_NAME)


def run_command_script(arguments: list[str]) -> int:
    """Run the command script with the specified command-line arguments."""
    command_script_path = get_command_script_path()
    completed_process = subprocess.run(
        [sys.executable, str(command_script_path), *arguments],
        check=False,
    )
    return completed_process.returncode


def main() -> int:
    """Forward the dropped Excel file name to the command script."""
    return run_command_script(sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
