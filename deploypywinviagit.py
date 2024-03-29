import os
import os.path as osp
import re
import sys
import stat
import shutil
import argparse
import subprocess
from pathlib import Path
from configparser import ConfigParser
from docutils.core import publish_file
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
    return replace_environment_variables(path.replace(match.group(0), env))


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
    system_ssh = osp.join(os.getenv('SystemRoot'), 'System32', 'OpenSSH', 'ssh.exe')
    assert osp.exists(system_ssh)
    subprocess.check_call(['git', 'clone', '--depth=1', src, dst], env=dict(os.environ,
        GIT_SSH=system_ssh))


def update_version_str(config: ConfigParser):
    if 'Repository' not in config.sections():
        raise RuntimeError('section `Repository` is required')
    if 'src' not in config['Repository']:
        raise RuntimeError('key `src` required in section Repository')
    if 'version' not in config['Repository']:
        return
    host, directory = config.get('Repository', 'src').split(':')
    dst = Path(replace_environment_variables(config.get('Repository', 'dst'))) / Path(config.get('Repository', 'version'))
    if not osp.exists(dst):
        raise RuntimeError('requested to write version to %s, but it doesnt exist' % dst)
    version = subprocess.check_output(['ssh', host, 'git', '-C', directory, 'describe', '--tags']).decode('ascii').strip()
    with open(dst, 'r', encoding='utf-8') as fin:
        contents = fin.read()
    version_written = False
    with open(dst, 'w', encoding='utf-8') as fout:
        for line in contents.splitlines():
            if line.startswith('__version__'):
                print('__version__ =', '"{:}"'.format(version), file=fout)
                version_written = True
            else:
                print(line.rstrip(), file=fout)
    if not version_written:
        print('WARNING version requested, but none found in', dst)
    print('version =', version)


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


def rst2html(source: Path):
    destination = source.parent / '{}.html'.format(source.stem)
    with open(source, 'r', encoding='utf8') as src:
        publish_file(source=src, destination_path=str(destination), writer_name='html')


def create_documentation_shortcut(name, path: Path):
    if not path.exists():
        raise RuntimeError('expecting documentation entry existing at %s' % path)
    desktop = Path(shell.SpecialFolders('Desktop'))
    assert desktop.exists()
    shortcut_path = desktop / '{:}.lnk'.format(name)
    print('creating', shortcut_path, '...')
    shortcut_obj = shell.CreateShortcut(str(shortcut_path))
    shortcut_obj.TargetPath = str(path)
    shortcut_obj.Save()


def create_documentation(config: ConfigParser):
    if 'Documentation' not in config.sections():
        print('No documentation requested')
        return
    if 'glob' not in config['Documentation']:
        raise RuntimeError('no files to generate documentation from specified via `glob=`')
    dst = Path(replace_environment_variables(config.get('Repository', 'dst')))
    assert dst.exists()
    source = None
    for source in dst.glob(config['Documentation']['glob']):
        print('rst2html', source)
        rst2html(source)
    if source is None:
        raise RuntimeError('no documentation source found in %s' % config['Documentation']['glob'])
    entry = config['Documentation'].get('desktop-entry')
    if entry:
        print('Requesting desktop entry to', source.parent / 'index.html')
        create_documentation_shortcut(entry, source.parent / 'index.html')
     

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
    python_exec = get_python_executable(development or verbose)
    if development or verbose:
        shortcut_obj.TargetPath = os.getenv('ComSpec')
        if verbose:
            args_tail = ' --debug'
        else:
            args_tail = ''
        if script:
            assert (dst / script).exists()
            args = '/K {:} {:}'.format(python_exec, Path(script))
        elif module:
            args = '/K {:} -m{:}'.format(python_exec, module)
        else:
            raise AssertionError()
        shortcut_obj.Arguments = args + args_tail
    else:
        shortcut_obj.TargetPath = python_exec
        if script:
            assert (dst / script).exists()
            shortcut_obj.Arguments = str(Path(script))
        elif module:
            shortcut_obj.Arguments = '-m' + module
        else:
            raise AssertionError()
    shortcut_obj.WorkingDirectory = str(dst)
    shortcut_obj.Save()


def identity_already_added(config: ConfigParser):
    src = config.get('Repository', 'src')
    user, _ = src.split('@')
    try:
        for line in subprocess.run(['ssh-add', '-l'], stderr=subprocess.PIPE).stderr.decode('ascii').splitlines():
            if line.startswith('Error connecting to agent'):
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                print('you have to start ssh-agent, e.g. in PowerShell:')
                print(" Set-Service -Name ssh-agent -StartupType 'Automatic'")
                print(" Start-Service ssh-agent")
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                raise SystemExit(2)
        for line in subprocess.check_output(['ssh-add', '-l'], stderr=subprocess.DEVNULL).decode('ascii').splitlines():
            if line.find(' {:}@'.format(user)):
                return True
    except subprocess.CalledProcessError:
        pass
    return False


def add_identity(identity_path):
    for _ in range(2):
        try:
            subprocess.check_call(['ssh-add', identity_path])
            print('ssh identity added to agent')
            break
        except subprocess.CalledProcessError:
            print('changing permissions for', identity_path)
        subprocess.check_call(['icacls', identity_path, '/inheritance:r'])
        subprocess.check_call(['icacls', identity_path, '/grant:r', '{:}:(R)'.format(os.getenv('USERNAME'))])


def openssh_available():
    try:
        subprocess.check_call(['where', 'ssh-add'], stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
        return True
    except subprocess.CalledProcessError:
        return False


def check_ssh_identity(config):
    identity_path = osp.abspath(osp.join(osp.dirname(__file__), 'ssh-identity'))
    if osp.exists(identity_path):
        if not openssh_available():
            print('install OpenSSH with the following PowerShell command:')
            print("  Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH*'")
            print("  Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0")
            print("  Set-Service -Name ssh-agent -StartupType 'Automatic'")
            print('and restart aftwerards')
            raise SystemExit(1)
        print('using', identity_path)
        if not identity_already_added(config):
            add_identity(identity_path)
    

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
    for encoding in ('utf-8', None):
        try:
            config.read(args.ini_file, encoding=encoding)
            break
        except UnicodeDecodeError:
            pass
    check_ssh_identity(config)
    clone_repository(config)
    update_version_str(config)
    create_desktop_entries(config, development=args.development, verbose=args.verbose)
    create_documentation(config)


if __name__ == '__main__':
    shell = Dispatch("WScript.Shell")
    main()

