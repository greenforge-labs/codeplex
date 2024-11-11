"""
Needs to be run as administrator!
Use codeplex.bat: Right click -> "Run as administrator"
"""

from pathlib import Path

from typing import TypeVar


def print_fail(msg: str):
    print(msg, end="\n\n")


def print_warning(msg: str):
    print(msg, end="\n\n")


def print_ok(msg: str):
    print(msg, end="\n\n")


T = TypeVar("T")


def select_option(options: list[T], *, none_msg: str, one_msg: str, many_msg: str) -> T:
    if len(options) < 1:
        print_fail(none_msg)
        exit(0)
    elif len(options) == 1:
        print_ok(one_msg.format(single_option=options[0]))
        return options[0]
    else:
        while True:
            print_ok(many_msg.format(num_options=len(options)))
            for i, option in enumerate(options):
                print(f"({i+1}) {str(option)}")
            selection = input("Selection: ")
            try:
                selection = int(selection) - 1
                if selection < 0:
                    raise ValueError()
                return options[selection]
            except (ValueError, IndexError):
                print_fail("Unknown selection!\n")


def find_codesys_install_paths() -> list[Path]:
    program_files = Path("C:/Program Files")
    program_files_x86 = Path("C:/Program Files (x86)")

    found_paths = []

    for prg_files_path in [program_files, program_files_x86]:
        for child in prg_files_path.iterdir():
            if child.is_dir() and "CODESYS" in child.name:
                found_paths.append(child)

    return found_paths


def find_codesys_sub_path(codesys_install_path: Path) -> Path:
    codesys_path = install_path / "CODESYS"

    if not codesys_path.exists() or not codesys_path.is_dir():
        print_fail(f"ERROR: expected directory to exist: {codesys_path}")
        exit(0) 
    
    return codesys_path


def get_duplicate_codesys_name() -> str:
    return input("Name for duplicate installation: ")


if __name__ == "__main__":
    try:
        print()
        
        install_paths = find_codesys_install_paths()
        install_path = select_option(
            install_paths,
            none_msg="No CODESYS install paths found!",
            one_msg="CODESYS install path found: {single_option}",
            many_msg="{num_options} CODESYS install paths found:",
        )

        codesys_path = find_codesys_sub_path(install_path)        

        duplicate_name = get_duplicate_codesys_name()

        # TODO
        # copy codesys path
        # find CODESYS Program Data path under codesys_path / Settings / RepositoryLocations.ini (read with configparser)
        # create new CODESYS Program Data path for duplicate
        # update RepositoryLocations.ini in duplicate to point at new Program Data path
        # create a desktop shortcut for the duplicate
        
    except PermissionError:
        print_fail(f"Permission error! Are you running as administrator?")
        exit(0)
