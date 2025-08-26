import pkg_resources
import sys

# Python standard library modules (common ones to ignore)
stdlibs = sys.builtin_module_names

installed = {pkg.key: pkg.version for pkg in pkg_resources.working_set}

print("📦 Installed third-party packages (excluding stdlib):\n")
for name, version in installed.items():
    print(f"{name}=={version}")
