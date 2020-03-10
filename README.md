# VSCode Portable

run vscode portable with custom envs and make it auto update

## Get started

1. create a folder
2. put VSCode.exe in it
3. run it then the program will download latest VSCode (win32-x64) and make it portable
4. edit the VSCode.ini to add you own envs

## Configuration

File directory structure example

```
VSCODE
│  VSCode.exe
│  VSCode.ini
│
├─Portable
│  ├─jdk
│  │  └─bin
│  └─node
└─VSCode
```

To add

- jdk→JAVA_HOME

- jdk\bin→PATH

- node→PATH

You just need to config PortableEnvs like this:

```ini
[Default]
CheckUpdateInterval=7
VSCodeBranch=win32-x64
PortableEnvs=Path:node<jdk\bin|JAVA_HOME:jdk
```

- use ‘|’ to split different envs name
- use ‘:’ to split env name and value
- use ‘<’ to split different env value
