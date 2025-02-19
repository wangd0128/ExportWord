import pkg_resources
import sys

required_packages = {
    'gradio',
    'pandas',
    'python-docx',
    'pywin32'
}

installed_packages = {pkg.key for pkg in pkg_resources.working_set}
missing_packages = required_packages - installed_packages

if missing_packages:
    print(f"缺少以下包：{', '.join(missing_packages)}")
    sys.exit(1)
else:
    print("所有必要的包都已安装")
    sys.exit(0) 