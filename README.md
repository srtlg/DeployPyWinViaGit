# Deploy Python Scripts via GIT

Maybe this is just useful for me: I'd like to use git for deployment of python
scripts to Windows machines. With the following caveats

* those Windows machines already have a full installation of python
* it might be necessary to make and track minor changes right away
* I'd like not to expose the detailed time of every commit 

Usage:

```commandline
python deploypywinviagit.py ../application/desktop-entries.ini
```

## Configuration files

All information are read from ini files, like

```ini
[DesktopEntry-1]
script=bin/application.py
name=application
icon=icons/application.ico

[DesktiopEntry-2]
module=bin.other_application
name=other application
icon=icons/other-application.ico

[Repository]
src=git-helper@development-machine.local:src/application
dst=$PUBLIC/application
```

