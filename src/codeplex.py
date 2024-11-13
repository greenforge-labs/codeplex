"""
Needs to be run as administrator!
Use codeplex.bat: Right click -> "Run as administrator"
"""

from pathlib import Path
import shutil
import configparser
import os
import subprocess

from typing import TypeVar, Callable


def create_shortcut(location: Path, target: Path):
    subprocess.run(["powershell", f"$s=(New-Object -COM WScript.Shell).CreateShortcut('{location}');$s.TargetPath='{target}';$s.Save()"])


class ExitException(Exception):
    pass


def print_fail(msg: str, exit_: bool = True):
    print(f"ERROR: {msg}", end="\n\n")
    if exit_:
        raise ExitException()


def print_warning(msg: str):
    print(f"WARNING: {msg}", end="\n\n")


def print_info(msg: str):
    print(msg, end="\n\n")


def get_input(prompt: str):
    in_ = input(prompt)
    print()  # newline for spacing
    return in_


def codeplex_name(name: str) -> str:
    return f"CODESYS CODEPLEX ({name})"


T = TypeVar("T")


def select_yes_no(prompt: str, default: bool = False) -> bool:
    in_ = input(f"{prompt} ({"Y" if default else "y"}/{"n" if default else "N"}): ")
    
    if in_.lower() == "y":
        return True
    elif in_.lower() == "n":
        return False
    else:
        return default


def select_option(options: list[T], *, none_msg: str, one_msg: str, many_msg: str) -> T:
    if len(options) < 1:
        print_fail(none_msg)
        raise ExitException  # print_fail does this - we also do it here just for typing purposes
    elif len(options) == 1:
        print_info(one_msg.format(single_option=options[0]))
        return options[0]
    else:
        while True:
            print_info(many_msg.format(num_options=len(options)))
            for i, option in enumerate(options):
                print(f"({i+1}) {str(option)}")
            selection = get_input("Selection: ")
            try:
                selection = int(selection) - 1
                if selection < 0:
                    raise ValueError()
                return options[selection]
            except (ValueError, IndexError):
                print_fail("Unknown selection!\n")
                raise ExitException  # print_fail does this - we also do it here just for typing purposes


def find_install_paths() -> list[Path]:
    program_files = Path("C:/Program Files")
    program_files_x86 = Path("C:/Program Files (x86)")

    found_paths = []

    for prg_files_path in [program_files, program_files_x86]:
        for child in prg_files_path.iterdir():
            if child.is_dir() and "CODESYS" in child.name:
                found_paths.append(child)

    return found_paths


def find_codesys_path(install_path: Path) -> Path:
    codesys_path = install_path / "CODESYS"

    if not codesys_path.exists() or not codesys_path.is_dir():
        print_fail(f"Expected directory to exist: {codesys_path}")
    
    return codesys_path


def get_duplicate_name() -> str:
    return get_input("Name for duplicate installation: ")


def get_duplicate_codesys_path(install_path: Path, duplicate_name: str) -> Path:
    return install_path / codeplex_name(duplicate_name)


def copy_duplicate_codesys(codesys_path: Path, duplicate_codesys_path: Path) -> Callable[[], None]:
    if duplicate_codesys_path.exists():
        print_fail(f"CODESYS copy already exists at {duplicate_codesys_path}")

    print_info(f"Copying {codesys_path} to {duplicate_codesys_path}")
    shutil.copytree(codesys_path, duplicate_codesys_path)

    def cleanup_func():
        print_info(f"Cleaning up {duplicate_codesys_path}")
        shutil.rmtree(duplicate_codesys_path)
    
    return cleanup_func


def find_settings_path(codesys_path: Path) -> Path:
    settings_path = codesys_path / "Settings" / "RepositoryLocations.ini"
    if not settings_path.exists():
        print_fail(f"Expected settings file at: {settings_path}")
    return settings_path


def find_program_data_path(codesys_path: Path) -> Path:
    settings_path = find_settings_path(codesys_path)
    
    config = configparser.ConfigParser()
    config.read(settings_path, encoding="utf-16")
    
    sections = config.sections()

    if len(sections) < 1:
        print_fail(f"No config sections found in settings file: {settings_path}")
    
    section = sections[-1]
    print_info(f"Selected config from profile: {section}")

    library_path = Path(config[section]["Managed Libraries"])

    return library_path.parent


def get_duplicate_program_data_path(program_data_path: Path, duplicate_name: str) -> Path:
    return program_data_path.parent / codeplex_name(duplicate_name)


def copy_duplicate_program_data(program_data_path: Path, duplicate_program_data_path: Path) -> Callable[[], None]:
    if duplicate_program_data_path.exists():
        print_fail(f"CODESYS data already exists at {duplicate_program_data_path}")

    print_info(f"Copying {program_data_path} to {duplicate_program_data_path}")
    shutil.copytree(program_data_path, duplicate_program_data_path)

    def cleanup_func():
        print_info(f"Cleaning up {duplicate_program_data_path}")
        shutil.rmtree(duplicate_program_data_path)
    
    return cleanup_func


def update_program_data_settings(codesys_path: Path, program_data_path: Path):
    settings_path = find_settings_path(codesys_path)

    print_info(f"Updating settings file: {settings_path}")

    old_program_data_path = find_program_data_path(codesys_path)

    with open(settings_path, "r", encoding="utf-16") as f:
        settings_data = f.read()

    settings_data = settings_data.replace(str(old_program_data_path), str(program_data_path))
    
    with open(settings_path, "w", encoding="utf-16") as f:
        f.write(settings_data)


def create_desktop_shortcut(codesys_path: Path, duplicate_name: str):
    print_info("Creating desktop shortcut")
    exe_path = codesys_path / "Common" / "CODESYS.exe"
    desktop_shortcut = Path(os.environ["USERPROFILE"]) / "Desktop" / f"CODESYS ({duplicate_name}).lnk"
    create_shortcut(desktop_shortcut, exe_path)


def create_start_menu_entry(codesys_path: Path, duplicate_name: str):
    print_info("Creating start menu entry")
    exe_path = codesys_path / "Common" / "CODESYS.exe"
    start_menu_entry = Path("C:/") / "ProgramData" / "Microsoft" / "Windows" / "Start Menu" / "Programs" / f"CODESYS ({duplicate_name}).lnk"
    create_shortcut(start_menu_entry, exe_path)


if __name__ == "__main__":

    cleanup: list[Callable[[], None]] = []

    try:
        try:
            print()
            
            install_paths = find_install_paths()
            install_path = select_option(
                install_paths,
                none_msg="No CODESYS install paths found!",
                one_msg="CODESYS install path found: {single_option}",
                many_msg="{num_options} CODESYS install paths found:",
            )

            codesys_path = find_codesys_path(install_path)        

            duplicate_name = get_duplicate_name()

            duplicate_codesys_path = get_duplicate_codesys_path(install_path, duplicate_name)

            program_data_path = find_program_data_path(codesys_path)

            duplicate_program_data_path = get_duplicate_program_data_path(program_data_path, duplicate_name)

            cleanup_func = copy_duplicate_codesys(codesys_path, duplicate_codesys_path)
            cleanup.append(cleanup_func)

            cleanup_func = copy_duplicate_program_data(program_data_path, duplicate_program_data_path)
            cleanup.append(cleanup_func)

            update_program_data_settings(duplicate_codesys_path, duplicate_program_data_path)

            should_create_desktop_shortcut = select_yes_no("Create desktop shortcut?", default=True)
            if should_create_desktop_shortcut:
                create_desktop_shortcut(duplicate_codesys_path, duplicate_name)

            should_create_start_menu_entry = select_yes_no("Create start menu entry?", default=True)
            if should_create_start_menu_entry:
                create_start_menu_entry(duplicate_codesys_path, duplicate_name)
            
        except PermissionError:
            print_fail(f"Permission error! Are you running as administrator?")
    
    except ExitException:
        for f in cleanup:
            f()
