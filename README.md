# Fullwidth-Indentor-Oneclick
一键全角空格缩进！

因为可能有人看，先写个最基础版本，以后再装修.jpg

### 依赖环境

---

仅当您使用纯脚本版时才需要看本节。可以使用Start.bat一键运行，也可以在powershell中输入以下指令：

建议在虚拟环境中运行（以Python 3.10+为例）：

```bash
python -m venv venv
venv\Scripts\activate
```
若觉得安装到全局没有问题，可忽略本步骤。
运行必须环境：

```bash
pip install PySide6 python-docx charset-normalizer
```

用于获取「快速访问」功能：

```bash
pip install pywin32
```

假如您不知道自己是否已安装pyinstaller也可安装：

```bash
pip install pyinstaller
```

以上均只需pip一次，之后直接 venv\Scripts\activate + python 脚本完整名称（带.py后缀）即可启动。
在安装到全局的情况下，无需再 venv\Scripts\activate ，在终端中直接python 脚本完整名称（带.py后缀）即可。
