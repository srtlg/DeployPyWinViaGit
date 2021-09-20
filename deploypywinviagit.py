import os
import re
import sys
import stat
import shutil
import argparse
import subprocess
from pathlib import Path
from configparser import ConfigParser
from win32com.client import Dispatch

var = re.compile(r'\$[A-Za-z][A-Za-z0-0]*')
shell = None


def replace_environment_variables(path: str):
    match = var.search(path)
    if match is None:
        return path
    sym = match.group(0)[1:]
    env = os.getenv(sym, None)
    if env is None:
        raise RuntimeError('Requested environment variable %s does not exist' % sym)
    return path.replace(match.group(0), env)


# https://stackoverflow.com/questions/1889597/deleting-read-only-directory-in-python/1889686#1889686
def remove_readonly(func, path, excinfo):
    os.chmod(path, stat.S_IWRITE)
    func(path)


def clone_repository(config: ConfigParser):
    if 'Repository' not in config.sections():
        raise RuntimeError('section `Repository` is required')
    if 'src' not in config['Repository']:
        raise RuntimeError('key `src` required in section Repository')
    if 'dst' not in config['Repository']:
        raise RuntimeError('key `dst` required in section Repository')
    src = config.get('Repository', 'src')
    dst = Path(replace_environment_variables(config.get('Repository', 'dst')))
    if dst.exists():
        shutil.rmtree(dst, onerror=remove_readonly)
    subprocess.check_call(['git', 'clone', '--depth=1', src, dst])


def get_python_executable(development):
    if development:
        return sys.executable
    elif sys.executable.endswith('python.exe'):
        python = sys.executable[:-len('python.exe')] + 'pythonw.exe'
        if Path(python).exists():
            return python
        else:
            return sys.executable
    else:
        return sys.executable


def create_desktop_entry(config: ConfigParser, section: str, development=False, verbose=False):
    dst = Path(replace_environment_variables(config.get('Repository', 'dst')))
    assert dst.exists()
    if 'name' not in config[section]:
        raise RuntimeError('key `name` required in section %s' % section)
    if 'icon' not in config[section]:
        raise RuntimeError('key `icon` required in section %s' % section)
    script = config.get(section, 'script', fallback=None)
    module = config.get(section, 'module', fallback=None)
    if script is None and module is None:
        raise RuntimeError('one key `script` or `module` required in section %s' % section)
    name = config.get(section, 'name')
    icon = config.get(section, 'icon')
    icon_path = dst / icon
    if not icon_path.exists():
        raise RuntimeError('icon of %s not found in %s' % (section, icon_path))
    desktop = Path(shell.SpecialFolders('Desktop'))
    assert desktop.exists()
    shortcut_path = desktop / '{:}.lnk'.format(name)
    print('creating', shortcut_path, '...')
    shortcut_obj = shell.CreateShortcut(str(shortcut_path))
    shortcut_obj.IconLocation = str(icon_path)
    if development or verbose:
        shortcut_obj.TargetPath = os.getenv('ComSpec')
        if verbose:
            args_tail = ' --debug'
        else:
            args_tail = ''
        if script:
            assert (dst / script).exists()
            args = '/K {:} {:}'.format(get_python_executable(development), Path(script))
        elif module:
            args = '/K {:} -m{:}'.format(get_python_executable(development), module)
        else:
            raise AssertionError()
        shortcut_obj.Arguments = args + args_tail
    else:
        shortcut_obj.TargetPath = get_python_executable(development)
        if script:
            assert (dst / script).exists()
            shortcut_obj.Arguments = str(Path(script))
        elif module:
            shortcut_obj.Arguments = '-m' + module
        else:
            raise AssertionError()
    shortcut_obj.WorkingDirectory = str(dst)
    shortcut_obj.Save()
    

def create_desktop_entries(config: ConfigParser, **kwargs):
    for section in [i for i in config.sections() if i.startswith('DesktopEntry')]:
        create_desktop_entry(config, section, **kwargs)


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('ini_file')
    p.add_argument('-d', '--development', action='store_true', 
                   help='create shortcuts that open a shell that stays open')
    p.add_argument('-v', '--verbose', action='store_true',
                   help='add --debug to the shortcut, implies --development')
    return p.parse_args()


def main():
    args = parse_args()
    config = ConfigParser()
    if not Path(args.ini_file).exists():
        raise RuntimeError('expecting an existing ini-file at `%s`' % args.ini_file)
    config.read(args.ini_file)
    clone_repository(config)
    create_desktop_entries(config, development=args.development, verbose=args.verbose)


if __name__ == '__main__':
    shell = Dispatch("WScript.Shell")
    main()

